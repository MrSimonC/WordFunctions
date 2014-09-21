#http://msdn.microsoft.com/en-us/library/ff837519(v=office.14).aspx
import types
import win32com.client as win32
from win32com.client import constants as c
import re

def openWordFile(file, visible=False):   #takes in a filename and returns a word document object
    word = win32.Dispatch('Word.Application')
    word.Visible = visible
    word.Documents.Open(file, False, True)    #FileName, ConfirmConversions, ReadOnly
    return word

def closeWordDocument(word):
    word.ActiveDocument.Close()

def cleanUnicode(unicodeText):
    return re.sub(r"[\r\n\t\x07\x0b]", "", unicodeText)

def printToNewDocument(content):    #not used - prints unicode value to a new word document
    word = win32.Dispatch('Word.Application')
    word.Documents.Add()
    docnew = word.ActiveDocument.Range(0, 0)
    docnew.InsertBefore(content)

def findPostitionInTable(doc, text):  #search all tables, return table, row, column containing text
    for t in range(1, doc.Tables.Count + 1):
        try:    #2 cells within a cell will produce errors (e.g. column count will be too high for the loop)
            for r in range(1, doc.Tables(t).Rows.Count + 1):
                try:
                    for c in range(1, doc.Tables(t).Columns.Count + 1):
                        header = doc.Tables(t).Cell(Row=r,Column=c).Range.Text
                        headerArray = ''.join(header).splitlines()  #split characters against unicode new lines
                        for item in headerArray:
                            if item == text:    #if you find your search term
                                return t, r, c  #return it's position (table, row, column)
                except:
                    c = c + 1   #loop out until a cell reference does exist
        except:
                r = r + 1

def findTableContent(word, search, ColumnOffset=0, RowOffset=0):    #search a document full of tables to return an offset cell value
    #try:
        doc = word.ActiveDocument
        tNumber, tRow, tColumn = findPostitionInTable(doc, search) #find ref to this text in the document
        content = doc.Tables(tNumber).Cell(Row=tRow + RowOffset, Column=tColumn + ColumnOffset).Range.Text
        return cleanUnicode(content)
    #except:
     #   return False
