#!python
import sys
import os
import uno
from typing import Any

XController = Any
XComponent = Any

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

def absoluteUrl(filename):
    """Constructs absolute path to the current dir in the format required by PyUNO that working with files"""
    mbPrefix = '' if filename[0] in '/~' else os.path.realpath('.') + '/'
    return 'file:///' + mbPrefix + os.path.expanduser(filename)

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
    row.getCellByPosition(3, 0).String = pivotRow.count

def emptyRows(slideRows, since):
    for i in range(since, slideRows.Count):
        row = slideRows.getByIndex(i)
        row.getCellByPosition(1, 0).String = ''
        row.getCellByPosition(2, 0).String = ''
        row.getCellByPosition(3, 0).String = ''

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

# courtesy to https://forum.openoffice.org/en/forum/viewtopic.php?f=20&t=63966
# BUG: this works, but LO sometimes too slow in doing the copy, so by time you
# execute Paste, slide wasn't copied yet, which gonna result in further problems.
def copySlideTo(srcApp: XComponent, dstApp: XComponent,
                slide, insert_after: int, smgr):
    srcController = srcApp.CurrentController
    dstController = dstApp.CurrentController
    dispatcher = smgr.createInstance("com.sun.star.frame.DispatchHelper")
    srcController.setCurrentPage(slide)
    ## begin: terrible magic to get damn thing copied
    dispatcher.executeDispatch(srcController.Frame, ".uno:DiaMode", "", 0, ())
    dispatcher.executeDispatch(srcController.Frame, ".uno:Copy", "", 0, ())
    dispatcher.executeDispatch(srcController.Frame, ".uno:NormalMultiPaneGUI", "", 0, ())
    ## end: terrible magic to get damn thing copied
    dstController.setCurrentPage(dstApp.DrawPages.getByIndex(insert_after))
    dispatcher.executeDispatch(dstController.Frame, ".uno:Paste", "", 0, ())
    return dstApp.DrawPages.getByIndex(insert_after+1)

# duplicates `slide`. Also sets the duplicated slide as "current".
def copySlide(drawController, slide, smgr):
    dispatcher = smgr.createInstance("com.sun.star.frame.DispatchHelper")
    drawController.setCurrentPage(slide)
    dispatcher.executeDispatch(drawController.Frame, ".uno:DuplicatePage", "", 0, ())
    return drawController.CurrentPage


def fillTailTables(tailTablesSlide, drawController, smgr, sheetRowsIter):
    tailTables = tablesFromSlide(tailTablesSlide)
    while sheetRowsIter != None:
        for i, table in enumerate(tailTables):
            sheetRowsIter = fillSlideTableFromSheet(table, sheetRowsIter)
            if sheetRowsIter == None:
                for iUnusedTables in range(i+1, len(tailTables)):
                    tailTables[iUnusedTables].dispose()
                break
        if sheetRowsIter != None:
            newTailTablesSlide = copySlide(drawController, tailTablesSlide, smgr)
            return fillTailTables(newTailTablesSlide, drawController, smgr, sheetRowsIter)

def exitIfWrongArgs():
    if len(sys.argv) != 4:
        print("Wrong number of arguments. Usage:\n" \
              "{}: <file_tables_sample> <file_spreadsheet> <file_dst_presentation> <impressions|publishers>".format(sys.argv[0]))
        exit(-1);

def openDocuments(desktop) -> (XComponent, XComponent, XComponent):
    # todo: open frames hidden, see https://forum.openoffice.org/en/forum/viewtopic.php?f=44&t=41379
    sampleApp      = desktop.loadComponentFromURL(absoluteUrl(sys.argv[1]) ,"_blank", 0, ())
    spreadsheetApp = desktop.loadComponentFromURL(absoluteUrl(sys.argv[2]) ,"_blank", 0, ())
    dstApp         = desktop.loadComponentFromURL(absoluteUrl(sys.argv[3]) ,"_blank", 0, ())
    return (sampleApp, spreadsheetApp, dstApp)

# script <file_tables_sample> <file_spreadsheet> <file_dst_presentation> <impressions|publishers>
def main():
    exitIfWrongArgs()
    (desktop, smgr) = connectToLO()

    (sampleApp, spreadsheetApp, dstApp) = openDocuments(desktop)
    mentionsSheet = spreadsheetApp.Sheets.getByName('QQ')
    # todo: check the impressions|publishers arg
    influencersSorted = collectInfluencers(mentionsSheet)
    # publishersSorted  = collectPublishers(mentionsSheet)

    topSlideSample  = sampleApp.DrawPages.getByIndex(0)
    tailSlideSample = sampleApp.DrawPages.getByIndex(1)

    # todo: ask oxana: how to determine where tables should be placed in? A cmd arg?
    dstSlideTop = copySlideTo(sampleApp, dstApp, topSlideSample, 0, smgr) # todo: 0 for testing
    sheetRowsIter = fillSlideTableFromSheet(dstSlideTop, iter(influencersSorted))
    dstSlideTail = copySlideTo(sampleApp, dstApp, tailSlideSample, 1, smgr) # todo: 1 for testing
    sheetRowsIter = fillTailTables(dstSlideTail, impressApp.CurrentController,
                                   smgr, sheetRowsIter)
    assert sheetRowsIter == None, "BUG: some rows in the sheet haven't been processed"

if __name__ == "__main__":
    main()
