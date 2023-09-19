# Python module to handle excel tables.
import openpyxl

# My custom modules
from Table import Table
from Formats import Formats
from TableBuilder import TableBuilder

allDictionary = {
    "Formatos": {
        "AQ": "Antique",
        "BA": "Baguete",
        "CA": "Carre",
        "CE": "Cela",
        "CL": "Cela Lateral",
        "CO": "Coração",
        "CZ": "Cruz",
        "DI": "Disco",
        "ET": "Estrela",
        "FL": "Flor",
        "FO": "Folha",
        "GC": "Gota Chata",
        "GO": "Gota",
        "GOMO": "Gomo",
        "HX": "Hexagono",
        "IG": "Igreja",
        "LE": "Leque",
        "LO": "Losango",
        "LU": "Lua",
        "MC": "Meio Coração",
        "NA": "Navete",
        "OC": "Octogonal",
        "OV": "Oval",
        "RD": "Redondo",
        "SA": "Saia",
        "TH": "Trilhão",
        "TR": "Triângulo",
        "TV": "Trevo",
        "TZ": "Trapézio",
    },
    "Pedras": {
        "ZA": "Ametista",
        "ZB": "Bege",
        "ZE": "Esmeralda",
        "ZF": "Fumê",
        "ZG": "Granada",
        "ZL": "Lilás",
        "ZM": "Amarela",
        "ZN": "Negra",
        "ZO": "Laranja",
        "ZR": "Rosa",
        "ZS": "Branca",
        "ZP": "Branca",
        "ZT": "Tanzanita",
        "ZV": "Verde",
        "ZZ": "Azul",     
        "ZVCL": "Verde Clara"   
    },
    "Lapidações": {
        "FC": "Facetado",
        "CB": "Cabochão",
        "BE": "Briolet",
        "CH": "Chapa",
        "MI": "Millenium",
        "IR": "Irregular",
        "FU": "Furado",
        "QD": "Quadrante",
        "/I": ""
    },
    "Extras": {
        "CQ": "Canto Quebrado",
        "CV": "Canto Vivo",
        "FL": "Furo Lateral",
        "FT": "Furo Topo",
        "FF": "Furo Frontal",
        "MF": "Meio Furo",
        "ML": "Meio Lado",
        "OP": "Opaca",
        "SP": "Superior",
        "TO": "Torto",
        "FM": "Fume",
        "Q": "Quebrada",
        "/I": ""
    }
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

variations = Formats(allDictionary)

tbuilder = TableBuilder(
    tableStructure, 
    int(input("Insira o último ID: ")),  
    None,
    None,
    input("Insira o nome da pedra mãe 'Exemplo: Zircônia de Primeira': "),
    input("Insira o código da pedra mãe 'Exemplo: ZP': "),
    input("Apenas de primeira? [True ou False]: ")
)

tbuilder.automaticBuild(oldTable, variations)