# Python module to handle excel tables.
import openpyxl

# This is where our table will be stored and dissected.
class Table:
    def __init__(self, directory):
        self.table = openpyxl.load_workbook(directory)
        self.activeSheet = self.table.active
        self.sheets = {}
        for worksheet in self.table.worksheets:
            self.sheets.update({worksheet.title: worksheet})

    def setActiveSheet(self, newSheet):
        try:
            self.activeSheet = self.sheets[newSheet]
        except:
            print("[ERROR] Unknow Sheet - Check if the sheet exists in the Table.")

    def getActiveSheetElement(self, element):
        elementDic = {"columns": self.activeSheet.iter_cols(values_only=True),
                       "rows": self.activeSheet.iter_rows(values_only=True)}
        try:
            templist = []
            for x in elementDic[element]:
                if not x:
                    continue
                templist.append(x)
        except:
            print("[ERROR] Unknow Element - Check if the passed property exists.")
        return templist

x = Table("testExcel.xlsx")
print(x.getActiveSheetElement("columns"))