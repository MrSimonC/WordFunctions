import win32com.client as win32
import re
__version__ = 1.1
# http://msdn.microsoft.com/en-us/library/ff837519(v=office.14).aspx


class Word:
    def __init__(self):
        self.word = win32.Dispatch('Word.Application')
        self.active_document = None

    def open_word_file(self, file, visible=False):   # takes in a filename and returns a word document object
        self.word.Visible = visible
        self.word.Documents.Open(file, False, True)    # FileName, ConfirmConversions, ReadOnly
        self.active_document = self.word.ActiveDocument

    def close(self):
        self.word.ActiveDocument.Close()

    @staticmethod
    def clean_unicode(unicode_text):
        return re.sub(r"[\r\n\t\x07\x0b]", "", unicode_text)

    def _find_position_in_table(self, search):
        """
        Search all tables, return table, row, column containing text
        :param search: text to search for
        :return: table, row, column
        """
        # Search all tables, return table, row, column containing search text
        column = 0
        row = 0
        for table in range(1, self.active_document.Tables.Count + 1):
            try:    # 2 cells within a cell will produce errors (e.g. column count will be too high for the loop)
                for row in range(1, self.active_document.Tables(table).Rows.Count + 1):
                    try:
                        for column in range(1, self.active_document.Tables(table).Columns.Count + 1):
                            header = self.active_document.Tables(table).Cell(Row=row,Column=column).Range.Text
                            header_array = ''.join(header).splitlines()
                            for item in header_array:
                                if item == search:
                                    return table, row, column
                    except:
                        column += 1   # loop out until a cell reference does exist
            except:
                    row += 1

    def find_table_content(self, search, column_offset=0, row_offset=0):
        # search a document full of tables to return an offset cell value
        table, row, column = self._find_position_in_table(search)
        content = self.active_document.Tables(table).Cell(row + row_offset, column + column_offset).Range.Text
        return self.clean_unicode(content)
