# Python module to handle excel tables.
import openpyxl

# This class gets all the information about the old table and the formats to build a new one, following certain patterns.
class TableBuilder:
    def __init__(self, defaultTableStructure:list, lastID:int, shortDescription, longDescription, lapidation:str, motherStoneName:str, motherStoneSignature:str):
        self.lastID = lastID
        self.newTableInfo = list()
        self.defaultTableStructure = defaultTableStructure
        self.shortDescription = shortDescription
        self.longDescription = longDescription
        self.lapidation = lapidation
        self.motherStoneName = motherStoneName
        self.motherStoneSignature = motherStoneSignature

    def automaticBuild(self, table, variations):
        self.filterTable(table, variations)
        self.buildNewTableInfo(variations)
        self.saveNewTable()

    # Filters the old table, adding all the formats and sizes into a class "Formats"
    def filterTable(self, table, variations):
        for row in table.getActiveSheetElement("rows"):
            stoneType = row[1][:2]
            formatKey = row[1][2:]
            formatID = row[0]
            formatSize = row[2]

            variations.appendFormat(formatKey)
            variations.appendInfo(formatKey, formatID, formatSize, stoneType)

    def buildNewTableInfo(self, variations):
        self.newTableInfo.append(self.defaultTableStructure)
        for formats in variations.getFormats():
        # Retrive all variations of the currently iterable Format, and repeat for each one of them.
            variationInfo = variations.getFormats()[formats]
            for info in variationInfo:
                self.lastID += 1
                self.newTableInfo.append([self.lastID, "variation", info[0], self.motherStoneName, 1, 0, "visible", self.shortDescription, 
                self.longDescription, None, None, "taxable", "parent", 1, None, None, 0, 0, 0, 0, 0, 0, 0, None, None, 0, None, None, None, 
                None, None, None, self.motherStoneSignature, None, None, None, None, None, 0, "Cor", info[2], 1, 1, "Formato", 
                formats, 1, 1, "Tamanho", info[1], 1, 1, "Lapidação", self.lapidation, 1, 1])

    def saveNewTable(self):
        newTable = openpyxl.Workbook()
        newTableSheet = newTable.active

        for row in self.newTableInfo:
            newTableSheet.append(row)

        newTable.save("Nova Tabela.xlsx")