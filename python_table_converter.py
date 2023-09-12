# Python module to handle excel tables.
import openpyxl

# This class is where the old table is dissected.
class Table:
    def __init__(self, directory:str):
        self.table = openpyxl.load_workbook(directory)
        self.activeSheet = self.table.active
        self.sheets = dict()
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
            templist = list()
            for x in elementDic[element]:
                # Filter to remove None elements.
                x = list(filter(lambda y: False if y is None else True, x))
                templist.append(x)
            # Filter to remove Empty elements and the Header.
            templist.pop(0)
            templist = list(filter(lambda y: False if len(y) == 0 else True, templist))
        except:
            print("[ERROR] Unknow Element - Check if the passed element exists.")
        return templist
    
# This class gets all the formats and sizes of the inputted table.
class Formats:
    def __init__(self, dictionary:dict):
        self.formats = dict()
        self.dictionary = dictionary

    def appendFormat(self, formatID):
        if self.dictionary[formatID] not in self.formats:
            self.formats[self.dictionary[formatID]] = list()

    def retrieveFormat(self, formatID):
        try:
            return self.formats[self.dictionary[formatID]]
        except:
            print("[ERROR] Format Unknow - Check if this format was writed correctly or if it exists in the format list")

    def deleteFormat(self, formatID):
        del self.formats[formatID]

    def appendInfo(self, formatID, id, size):
        self.formats[self.dictionary[formatID]].append([id, size])

    def getFormats(self):
        return self.formats

# This class gets all the information about the old table and the formats to build a new one, following certain patterns.
class TableBuilder:
    def __init__(self, defaultTableStructure:list, lastID:int, motherStone:str, motherStoneSignature:str, shortDescription, longDescription, lapidation:str):
        self.lastID = lastID
        self.newTableInfo = list()
        self.defaultTableStructure = defaultTableStructure
        self.motherStone = motherStone
        self.motherStoneSignature = motherStoneSignature
        self.shortDescription = shortDescription
        self.longDescription = longDescription
        self.lapidation = lapidation

    def automaticBuild(self, table, variations):
        self.filterTable(table, variations)
        self.buildNewTableInfo(variations)
        self.saveNewTable()

    # Filters the old table, adding all the formats and sizes into a class "Formats"
    def filterTable(self, table, variations):
        for row in table.getActiveSheetElement("rows"):
            formatKey = row[1][2:]
            formatID = row[0]
            formatSize = row[2]

            variations.appendFormat(formatKey)
            variations.appendInfo(formatKey, formatID, formatSize)

    def buildNewTableInfo(self, variations):
        self.newTableInfo.append(self.defaultTableStructure)
        for formats in variations.getFormats():
        # Retrive all variations of the currently iterable Format, and repeat for each one of them.
            variationInfo = variations.getFormats()[formats]
            for info in variationInfo:
                self.lastID += 1
                self.newTableInfo.append([self.lastID, "variation", info[0], self.motherStone, 1, 0, "visible", self.shortDescription, 
                self.longDescription, None, None, "taxable", "parent", 1, None, None, 0, 0, 0, 0, 0, 0, 0, None, None, 0, None, None, None, 
                None, None, None, self.motherStoneSignature, None, None, None, None, None, 0, "Formato", formats, 1, 1, "Lapidação", 
                self.lapidation, 1, 1, "Tamanho", info[1], 1, 1])

    def saveNewTable(self):
        newTable = openpyxl.Workbook()
        newTableSheet = newTable.active

        for row in self.newTableInfo:
            newTableSheet.append(row)

        newTable.save("newTable.xlsx")

formatsDictionary = {
    "ZP": "Branca",
    "BA": "Baguete",
    "CA": "Carre",
    "CO": "Coração",
    "GO": "Gota",
    "NA": "Navete",
    "OC": "Octogonal",
    "OV": "Oval",
    "RD": "Redondo",
    "TR": "Triângulo",
    "TZ": "Trapézio",
    "CAEX": "Carre Extra",
    "RDEX": "Redonda Extra",
    "RDAZ": "Redonda Azul",
    "RDFM": "Redonda Fume",
    "RD-AAA": "Redonda Milheiro AAA",
    "RDM-RE": "Redonda Milheiro Rema",
    "RDM-EX": "Redonda Milheiro Extra",
    "RDM/TS": "Redonda Milheiros TS",
    "RDMIL": "Redonda Milheiro",
    "RDMIL*": "Redonda Milheiro Asterisco"
}
tableStructure = ([
    "ID", "Tipo", "SKU", "Nome", "Publicado", "Em Destaque?", "Visibilidade no catálogo", "Descrição curta", "Descrição", 
    "Data de preço promocional","Data de preço promocional", "Status da taxa", "Classe de taxa", "Em estoque?", "Estoque", 
    "Quantidade baixa de estoque", "São permitidas encomendas", "Vendido individualmente", "Peso (kg)", "Comprimento (cm)", 
    "Largura (cm)", "Altura (cm)", "Permitir avaliações", "Nota da compra", "Preço promocional", "Preço", "Categorias", "Tags", 
    "Classe de entrega", "Imagens", "Limite de download", "Dias para expirar o download", "Produto ascendente", "Grupo de produtos",
    "Aumentar vendas","Venda cruzada", "URL externa", "Texto do botão","Posição", "Nome do atributo 1", "Valores do atributo 1",
    "Visibilidade do atributo 1","Atributo global 1", "Nome do atributo 2", "Valores do atributo 2", "Visibilidade do atributo 2",
    "Atributo global 2", "Nome do atributo 3","Valores do atributo 3","Visibilidade do atributo 3", "Atributo global 3",
    "Nome do atributo 4", "Valores do atributo 4", "Visibilidade do atributo 4", "Atributo global 4"])

oldTable = Table(input("Insira o nome do arquivo da tabela: "))

variations = Formats(formatsDictionary)

tbuilder = TableBuilder(
    tableStructure, 
    int(input("Insira o último ID: ")),  
    None,
    None,
    "Facetado", 
    input("Insira o nome da pedra mãe 'Exemplo: Zircônia de Primeira': "),
    input("Insira o código da pedra mãe 'Exemplo: ZP': ")
)
tbuilder.automaticBuild(oldTable, variations)