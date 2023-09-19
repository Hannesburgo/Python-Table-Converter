# Python module to handle excel tables.
import openpyxl

# This class gets all the information about the old table and the formats to build a new one, following certain patterns.
class TableBuilder:
    def __init__(self, defaultTableStructure:list, lastID:int, shortDescription, longDescription, motherStoneName:str, motherStoneSignature:str, isFirstClass):
        self.lastID = lastID
        self.newTableInfo = list()
        self.defaultTableStructure = defaultTableStructure
        self.shortDescription = shortDescription
        self.longDescription = longDescription
        self.motherStoneName = motherStoneName
        self.motherStoneSignature = motherStoneSignature
        self.isFirstClass = eval(isFirstClass)

    def automaticBuild(self, table, variations):
        self.filterTable(table, variations)
        self.buildNewTableInfo(variations)
        self.saveNewTable()

    # Filters the old table, adding all the formats and sizes into a class "Formats"
    def filterTable(self, table, variations):
        rowInfo = list()
        for row in table.getActiveSheetElement("rows"):
            stoneId = row[0]
            stoneType = row[1][:2]
            stoneFormat = row[1][2:4]
            stoneSize = row[2]
            stoneExtraOne = row[1][4:6]
            stoneExtraTwo = row[1][6:]

            if stoneFormat == "OP":
                stoneFormat = stoneExtraOne
                stoneExtraOne = "FC"
                stoneExtraTwo = "OP"
            
            if not stoneExtraOne:
                stoneExtraOne = "FC"

            # Check VERDE CLARAR zircons
            if stoneType == "ZV" and stoneFormat == "CL":
                stoneFormat = stoneExtraOne
                stoneType = "ZVCL"
                if stoneExtraTwo:
                    stoneExtraOne = stoneExtraTwo
                    stoneExtraTwo = ""
                else:
                    stoneExtraOne = "FC"

            # Check GOMO zircons
            if stoneType == "GO" and stoneExtraOne == "MO":
                stoneFormat = "GOMO"
                stoneExtraOne = ""

            variations.appendFormat(stoneFormat)
            variations.appendInfo(stoneId, stoneType, stoneFormat, stoneSize, stoneExtraOne, stoneExtraTwo, self.isFirstClass)

    def buildNewTableInfo(self, variations):
        self.newTableInfo.append(self.defaultTableStructure)
        for formats in variations.getFormats():
        # Retrive all variations of the currently iterable Format, and repeat for each one of them.
            variationInfo = variations.getFormats()[formats]
            for info in variationInfo:
                self.lastID += 1
                self.newTableInfo.append([self.lastID, "variation", info[0], self.motherStoneName, 1, 0, "visible", self.shortDescription, 
                self.longDescription, None, None, "taxable", "parent", 1, None, None, 0, 0, 0, 0, 0, 0, 0, None, None, 0, None, None, None, 
                None, None, None, self.motherStoneSignature, None, None, None, None, None, 0, "Cor", info[1], 1, 1, "Formato", 
                formats, 1, 1, "Tamanho", info[2], 1, 1, "Lapidação", info[3], 1, 1])

    def saveNewTable(self):
        newTable = openpyxl.Workbook()
        newTableSheet = newTable.active

        for row in self.newTableInfo:
            newTableSheet.append(row)

        newTable.save("Nova Tabela DE PRIMEIRA.xlsx")