#!/usr/bin/python3
import uno
import os

# that's like a consts, but they isn't since consts not allowed in python 😝
HEADING1    = "Heading 1"
HEADING2    = "Heading 2"
HEADING5    = "Heading 5" #appendinx
STYLESFILE  = "./styles.odt"
PAGE_BEFORE = 4 #that is from enumeration — that awful API have problems with them

def rmLastEmptyLine(paragraph):
    cursor = paragraph.Text.createTextCursor() #it is created at start of document, not paragraph
    cursor.gotoRange(paragraph.End, False) #get end of the current paragraph
    cursor.gotoPreviousWord(True)
    if cursor.String.endswith('\n'):
        cursor.String = cursor.String[:-1] #strip last newline

def overwriteStyles(document, fromFile):
    """First arg is the document itself, the second is a string with path to style file in UNO API format"""
    styles = document.StyleFamilies
    styles.loadStylesFromURL(fromFile, styles.StyleLoaderOptions)

def addNumberingSomeHeading1n2s(text, h1TillBiblio=None, h1n2Numbered=None):
    enumeration = text.createEnumeration()
    heading1s = 0 #«2» is intro; «5» is 3-rd chapter; «7» is biblio
    while enumeration.hasMoreElements():
        par = enumeration.nextElement()
        if ( par.supportsService("com.sun.star.text.Paragraph") and
            par.ParaStyleName != None):
            if par.ParaStyleName == HEADING1:
                heading1s = heading1s + 1
                if ( heading1s <= 6 #till, but excluding biblio
                     and h1TillBiblio != None):
                    h1TillBiblio(par)
            if (par.ParaStyleName == HEADING1 or
                par.ParaStyleName == HEADING2):
                if (heading1s > 5 or heading1s == 1 or heading1s == 2):
                    par.setPropertyValue("NumberingStyleName", "NONE")
                else:
                    par.setPropertyValue("NumberingStyleName", "Numbering 1")
                    if h1n2Numbered != None:
                        h1n2Numbered(par)
                if heading1s >= 8: #fuck this shit
                    par.setPropertyValue("BreakType", 0) #remove page break
            if (heading1s == 7 and par.ParaStyleName == HEADING5): #just after the biblio, the first appendix
                par.setPropertyValue("BreakType", PAGE_BEFORE)


def insertNewlineAfterPar(par):
    cursor = par.Text.createTextCursor() #it is created at start of document, not paragraph!
    cursor.gotoRange(par.End, False) #get end of the current paragraph
    #cursor.goRight(1, False)
    cursor.String = "\n"

def insertSpaceStartPar(par):
    cursor = par.Text.createTextCursor() #it is created at start of document, not paragraph!
    cursor.gotoRange(par.Start, False) #get start of the current paragraph
    cursor.String = " "

def absoluteUrl(relativeFile):
    """Constructs absolute path to the current dir in the format required by PyUNO that working with files"""
    return "file:///" + os.path.realpath(".") + "/" + relativeFile

def parBreak(document, cursor):
    """Inserts a paragraph break at cursor position"""
    document.Text.insertControlCharacter( cursor.End,
                                          uno.getConstantByName('com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK'),
                                          False)

def updateTOC(document):
    """Update links in table of contents"""
    iList = document.getDocumentIndexes()
    assert(iList.getCount() > 0), "Table of contents not found, aborting!"
    for i in range(0, iList.getCount()):
        iList.getByIndex(i).update()

def setFirstPage(file):
    """Sets the first page style to «First Page»"""
    enumeration = file.Text.createEnumeration()
    enumeration.nextElement().PageDescName = 'First Page'

#connect to office, and get the file object
localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", localContext )
smgr = resolver.resolve( "uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager" )
remoteContext = smgr.getPropertyValue( "DefaultContext" )
desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",remoteContext)
file = desktop.loadComponentFromURL(absoluteUrl("./output/output.odt") ,"_blank", 0, ())

overwriteStyles(file, absoluteUrl(STYLESFILE))
#insert at the beginning the file with Chapter1
cursor = file.Text.createTextCursor()
cursor.gotoStart(False)
#to not screw the first paragraph formatting let's first make an empty one here
parBreak(file, cursor)
cursor.gotoStart(False)
cursor.insertDocumentFromURL(absoluteUrl("./Ch1.odt"), ())
addNumberingSomeHeading1n2s(file.Text, insertNewlineAfterPar, insertSpaceStartPar)
updateTOC(file) #the inserted TOC needs to be updated

#save the file
file.storeAsURL(absoluteUrl("./output/Курсовая.doc"),())
file.storeAsURL(absoluteUrl("./output/Курсовая.odt"),())
file.dispose()

print("Not implemented: removing Appendinx entries from the end of TOC")