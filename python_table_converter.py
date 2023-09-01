# Python module to handle excel tables.
import openpyxl

# This is where the table will be stored and dissected.
class Table:
    def __init__(self, directory):
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

# This is where all stones are going to be stored and checked
class Formats:
    def __init__(self, dictionary):
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

formatsDictionary = {
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

oldTable = Table("LISTCUS-zp a.xlsx")
variations = Formats(formatsDictionary)

# Filters the old table, adding all the formats and sizes into a class "Formats"
def filterOldTable():
    for row in oldTable.getActiveSheetElement("rows"):
        formatKey = row[1][2:]
        formatID = row[0]
        formatSize = row[2]

        variations.appendFormat(formatKey)
        variations.appendInfo(formatKey, formatID, formatSize)
    # print(variations.getFormats())

defaultTableStructure = ([
    "ID", "Tipo", "SKU", "Nome", "Publicado", "Em Destaque?", "Visibilidade no catálogo", "Descrição curta", "Descrição", 
    "Data de preço promocional","Data de preço promocional", "Status da taxa", "Classe de taxa", "Em estoque?", "Estoque", 
    "Quantidade baixa de ???", "São permitidas encomendas", "Vendido individualmente", "Peso (kg)", "Comprimento (cm)", 
    "Largura (cm)", "Altura (cm)", "Permitir avaliações", "Nota da compra", "Preço promocional", "Preço", "Categorias", "Tags", 
    "Classe de entrega", "Imagens", "Limite de download", "Dias para expirar o ???", "Produto ascendente", "Grupo de produtos",
    "Aumentar vendas","Venda cruzada", "URL externa", "Texto do botão","Posição", "Nome do atributo 1", "Valores do atributo 1",
    "Visibilidade do atributo 1","Atributo global 1", "Nome do atributo 2", "Valores do atributo 2", "Visibilidade do atributo 2",
    "Atributo global 2", "Nome do atributo 3","Valores do atributo 3","Visibilidade do atributo 3", "Atributo global 3",
    "Atributo padrão 1"])
newTableInfo = list()

motherStone = "Zircônia de Primeira"
motherStoneSignature = "ZP"
shortDescription = None
longDescription = None
lapidation = "Facetado"

def buildnewTableInfo(header):
    lastID = 14206

    newTableInfo.append(header)
    for formats in variations.getFormats():
        # Retrive all variations of the currently iterable Format, and repeat for each one of them.
        variationInfo = variations.getFormats()[formats]
        for info in variationInfo:
            lastID += 1
            newTableInfo.append([lastID, "variation", info[0], motherStone, 1, 0, "visible", shortDescription, longDescription, None, None, 
                "taxable", "parent", 1, None, None, 0, 0, 0, 0, 0, 0, 0, None, None, 0, None, None, None, None, None, None, motherStoneSignature, 
                None, None, None, None, None, 0, "Formato", formats, None, 1, "Lapidação", lapidation, None, 1, "Tamanho", info[1], None, 1])

def buildHeaderProductAtt():
    mainProductAtt = ([], [])
    for formats in variations.getFormats():
        mainProductAtt[0].append(formats)
        variationInfo = variations.getFormats()[formats]
        for info in variationInfo:
            mainProductAtt[1].append(info[1])
    
    tempForm = ','.join(map(str, mainProductAtt[0]))
    tempSizes = ','.join(map(str, mainProductAtt[1]))
    print(tempForm)
    print("\n")
    print(tempSizes)

def saveNewTable():
    newTable = openpyxl.Workbook()
    newTableSheet = newTable.active

    for row in newTableInfo:
        newTableSheet.append(row)

    newTable.save("newTable.xlsx")

filterOldTable()
buildnewTableInfo(defaultTableStructure)
buildHeaderProductAtt()
      