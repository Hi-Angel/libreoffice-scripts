#!python
import uno

class PivotRow:
    def __init__(self, author, views, count):
        self.author = author
        self.views  = views
        self.count  = count

    def __str__(self):
        return "PivotRow {{author = {}, views = {}, count = {}}}" \
            .format(self.author, self.views, self.count)

    @classmethod
    def fromRow(self, row):
        return self(row.getCellByPosition(0, 0).String,
                    row.getCellByPosition(1, 0).String,
                    row.getCellByPosition(2, 0).String)

# run libreoffice as:
# soffice --calc --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"
def connectToLO():
    # get the uno component context from the PyUNO runtime
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext )
    # connect to the running office
    ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    # return (desktop, localContext)
    return (desktop, smgr)

def absoluteUrl(relativeFile):
    """Constructs absolute path to the current dir in the format required by PyUNO that working with files"""
    return "file:///" + os.path.realpath(".") + "/" + relativeFile

def getLOInstances(desktop):
    loInstances = [c for c in desktop.Components]
    assert(len(loInstances) == 2)
    CALC_STR    = 'ScModelObj'
    IMPRESS_STR = 'SdXImpressDocument'
    if IMPRESS_STR not in loInstances[0].ImplementationName:
        loInstances = [loInstances[1], loInstances[0]]
    assert(CALC_STR in loInstances[1].ImplementationName)
    return loInstances

def tablesFromSlide(slide):
    return [item for item in slide if item.ShapeType == 'com.sun.star.drawing.TableShape']

def collectNamesImpress(slides):
    names = {} # String : Row
    for slide in slides:
        tables = tablesFromSlide(slide)
        for table in tables:
            for row in table.Model.Rows:
                names[row.getCellByPosition(1, 0).String] = row
    return names

def dropFilters(filters):
    for filter in filters:
        filter.IsHidden = True

def setFilters(filters, filtersToSet):
    # that sucks, but if you set IsHidden to the same value it has, it still tries to
    # filter stuff out, bogging CPU. So below I check before altering IsHidden.
    namesToFilters = {}
    for filter in filters:
        namesToFilters[filter.Name] = filter
    for wantedFilterNames in filtersToSet:
        filter = namesToFilters.pop(wantedFilterNames)
        if filter.IsHidden != False:
            filter.IsHidden = False
    for _,filter in namesToFilters.items():
        if filter.IsHidden != True:
            filter.IsHidden = True

# OutputRange of pivot table includes "technical rows" at the beginning and end
def pivotTableUsedRangeMentions(pivotTable, sheet):
    N_TECHNICAL_ROWS_AT_START = 5
    N_TECHNICAL_ROWS_AT_END   = 1
    range_raw = pivotTable.OutputRange
    return sheet.getCellRangeByPosition(range_raw.StartColumn,
                                        range_raw.StartRow + N_TECHNICAL_ROWS_AT_START,
                                        range_raw.EndColumn,
                                        range_raw.EndRow   - N_TECHNICAL_ROWS_AT_END)

# Rows -> Int -> [String]
def rowToStrings(row, ncols):
    ret = []
    for iCol in range(0, ncols):
        ret.append(row.getCellByPosition(iCol, 0).String)
    return ret

# Sheet -> [String] -> [(int, Row)]; where Int is a number of views, and Row is the
# row
def collectFromPivotTable(sheet, filterNames):
    pilotTable = sheet.DataPilotTables.getByIndex(0)
    filters = pilotTable.DataPilotFields.getByName('Author type').Items
    # "influencers" means bloggers and celebrity
    setFilters(filters, filterNames)
    ret = []
    pivotRange = pivotTableUsedRangeMentions(pilotTable, sheet)
    for row in pivotRange.Rows:
        pivotRow = PivotRow.fromRow(row)
        mb_views = pivotRow.views # aka impressions
        ret.append((int(mb_views.partition(',')[0]) if mb_views else 0, # 0 is required for sorting
                    pivotRow))
    ret.sort(key = lambda pair: pair[0], reverse = True) # most views first
    return ret

# [(Int, Row)],
def collectPublishers(sheet):
    return collectFromPivotTable(sheet, ['Publisher'])

# [(Int, Row)],
def collectInfluencers(sheet):
    return collectFromPivotTable(sheet, ['Blogger', 'Celebrity'])

# Row -> PivotRow -> ()
def fillSlideRow(row, pivotRow, views):
    row.getCellByPosition(1, 0).String = pivotRow.author
    viewsStr = str(int(views / 1000)) + ' K' if views > 1000 else str(views)
    row.getCellByPosition(2, 0).String = viewsStr

def emptyRows(slideRows, since):
    for i in range(since, slideRows.Count):
        row = slideRows.getByIndex(i)
        row.getCellByPosition(1, 0).String = ''
        row.getCellByPosition(2, 0).String = ''

# SlideTable -> Iter (Int, PivotRow) -> Iter (Int, Rows)
def fillSlideTableFromSheet(slideTable, sheetRowsIter):
    nSlideRowsPassed = 0
    slideRows = slideTable.Model.Rows
    for (i, oneSlideRow), (views, oneSheetRow) in zip(enumerate(slideRows),
                                                      sheetRowsIter):
        nSlideRowsPassed += 1
        if i == 0:
            continue # 1-st slide-table row is a header
        fillSlideRow(oneSlideRow, oneSheetRow, views)
    if nSlideRowsPassed < slideTable.Model.Rows.Count: # not enough pivot results
        emptyRows(slideRows, nSlideRowsPassed)
        return None
    else: # means the iter has more elements
        return sheetRowsIter

# duplicates `slide`. Also sets the duplicated slide as "current".
def copySlide(drawController, slide, smgr):
    dispatcher = smgr.createInstance("com.sun.star.frame.DispatchHelper")
    drawController.setCurrentPage(slide)
    dispatcher.executeDispatch(drawController.Frame, ".uno:DuplicatePage", "", 0, ())
    return drawController.CurrentPage


def fillTailTables(tailTables, tailTablesSlide, drawController, smgr, sheetRowsIter):
    while sheetRowsIter != None:
        for i, table in enumerate(tailTables):
            sheetRowsIter = fillSlideTableFromSheet(table, sheetRowsIter)
            if sheetRowsIter == None:
                for iUnusedTables in range(i+1, len(tailTables)):
                    tailTables[iUnusedTables].dispose()
                break
        if sheetRowsIter != None:
            newTailTablesSlide = copySlide(drawController, tailTablesSlide, smgr)
            newTailTables = tablesFromSlide(newTailTablesSlide)
            return fillTailTables(newTailTables, newTailTablesSlide, drawController, smgr, sheetRowsIter)

(desktop, smgr) = connectToLO()

# todo: this is probably slides sample, not a generic "draw app"?
(impressApp, spreadsheetApp) = getLOInstances(desktop)
mentionsSheet = spreadsheetApp.Sheets.getByName('QQ')
influencersSorted = collectInfluencers(mentionsSheet)
# publishersSorted  = collectPublishers(mentionsSheet)

topSlideSample   = impressApp.DrawPages.getByIndex(0)
topTableSample   = tablesFromSlide(topSlideSample)[0]
tailSlideSample  = impressApp.DrawPages.getByIndex(1)
tailTablesSample = tablesFromSlide(tailSlideSample) # left table

sheetRowsIter = fillSlideTableFromSheet(topTableSample, iter(influencersSorted))
sheetRowsIter = fillTailTables(tailTablesSample, tailSlideSample,
                               impressApp.CurrentController, smgr, sheetRowsIter)
assert sheetRowsIter == None, "BUG: some rows in the sheet haven't been processed"

# impressApp.storeAsURL(absoluteUrl("output.odp"),())
# impressApp.dispose()
