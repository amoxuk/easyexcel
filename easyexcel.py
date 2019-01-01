import win32com.client

class ExcelApp:
    COM_NAME = "Excel.Application"
    def __init__(self):
        self._xlApp = win32com.client.Dispatch("Excel.Application")
        self._xlApp.Visible = False
        self._xlApp.DisplayAlerts = 0

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self._xlApp.Quit()

    def openBook(self, excelFile, password=None):
        if password:
            return ExcelBook(self._xlApp.Workbooks.Open(excelFile,
                0, True, 3, password))
        else:
            return ExcelBook(self._xlApp.Workbooks.Open(excelFile))

class ExcelBook:
    def __init__(self, excelBook):
        self._xlBook = excelBook
        self.name = self._xlBook.Name

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()

    def close(self, isSave=0):
        self._xlBook.Close(SaveChanges=isSave)

    def openSheet(self, sheetName):
        return ExcelSheet(self._xlBook.Worksheets(sheetName))

    def getSheet(self, index):
        count = self._xlBook.Sheets.Count
        if (index >= count):
            raise IndexError("index = " + str(index) + ", len = " + str(count))
        return ExcelSheet(self._xlBook.Sheets.Item(index))

class ExcelSheet:
    def __init__(self, excelSheet):
        self._xlSheet = excelSheet
        self.name = self._xlSheet.name

    def get(self, x, y):
        if (x >= self.getRowCount() or y >= self.getColumnCount()):
            raise IndexError("x = " + str(x) + " < " + str(self.getRowCount())
                             + " , y = " + str(y) + " < " + str(self.getColumnCount()))
        return self._xlSheet.Cells(x + 1, y + 1).Text

    def getColumnCount(self):
        return self._xlSheet.usedRange.Columns.Count

    def getRowCount(self):
        return self._xlSheet.usedRange.Rows.Count

    def asArray(self):
        columns = self.getColumnCount()
        rows = self.getRowCount()
        data = [[None for col in range(columns)] for row in range(rows)]
        for i in range(rows):
            for j in range(columns):
                data[i][j] = self.get(i, j)
        return data

if __name__ == "__main__":
    import re
    import os

    WORK_PATH = os.getcwd()+os.path.sep

    with ExcelApp() as app:
        for fname in os.listdir(WORK_PATH):
            if re.match("^[\d|\D]*.xlsx?$", fname):
                with app.openBook(WORK_PATH + fname) as book:
                    firstSheetName = book.getSheet(1).name
                    sheet = book.openSheet(firstSheetName)
                    print("[BookName=" + book.name + "][SheetName=" + sheet.name + "]:")
                    columns = sheet.getColumnCount()
                    rows = sheet.getRowCount()
                    print("Size: " + str(rows) + " * " + str(columns))
                    print(sheet.asArray())
                    print(type(app._xlApp))
