# Python module to handle excel tables.
import openpyxl

# My custom modules
from Table import Table
from Formats import Formats
from TableBuilder import TableBuilder

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