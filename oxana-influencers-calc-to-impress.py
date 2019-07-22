#!python
import uno

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
    return desktop

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
                                        range_raw.EndRow - N_TECHNICAL_ROWS_AT_END)

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
        mb_views = row.getCellByPosition(1, 0).String # aka impressions
        ret.append((int(mb_views.partition(',')[0]) if mb_views else 0, # 0 is required for sorting
                    rowToStrings(row, pivotRange.Columns.Count)))
    ret.sort(key = lambda pair: pair[0], reverse = True) # most views first
    return ret

# [(Int, Row)],
def collectPublishers(sheet):
    return collectFromPivotTable(sheet, ['Publisher'])

# [(Int, Row)],
def collectInfluencers(sheet):
    return collectFromPivotTable(sheet, ['Blogger', 'Celebrity'])

desktop = connectToLO()
(drawApp, spreadsheetApp) = getLOInstances(desktop)
mentionsSheet = spreadsheetApp.Sheets.getByName('QQ')
influencersSorted = collectInfluencers(mentionsSheet)
publishersSorted  = collectPublishers(mentionsSheet)

# test code below
slide1 = drawApp.DrawPages.getByIndex(0)
table1 = tablesFromSlide(slide1)[0]
index = 0
for row in table1.Model.Rows:
    if index == 0:
        index += 1
        continue
    if index >= len(influencersSorted):
        break
    row.getCellByPosition(1, 0).String = influencersSorted[index][1][0] # name
    views = influencersSorted[index][0]
    views_str = str(int(views / 1000)) + ' K' if views > 1000 else str(views)
    row.getCellByPosition(2, 0).String = views_str
    index += 1


# fillTable(influencersSorted, getInflTable(mentionsSheet))
# fillTable(publishersSorted, getPublTable(mentionsSheet))
