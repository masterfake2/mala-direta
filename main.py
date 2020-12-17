import pandas as pd
from docxtpl import DocxTemplate

excel_path = r"./dataset/base.xls"
dataset = pd.read_excel(excel_path)

# Acha a linha que representa o CPF digitado

try:
	# Faz a busca no dataset e retorna a linha
	row = dataset.loc[dataset["CPF DEVEDOR"] == "001.441.051-60"]
except Exception as e:
	# Linha não foi encontrada ou seja registro não existe, ou chave não existe
	row = False
	pass
row = dataset.loc[dataset["CPF DEVEDOR"] == "001.441.051-60"]

if(row is not False):
	print("Debug")
	#dados devedor
	nome = row["NOME"][0]
	cpf  = row["CPF DEVEDOR"][0]
	unidade = row["UNIDADE"][0]
	residencial = row["RESIDENCIAL"][0]

	#dados credor
	credor = ""
	cnpj   = ""

	#divida
	valor = ""
	data_entrada = ""
	numero_parcelas = ""
	valor_parcela = ""
	primeiro_pagamento = ""
	ultimo_pagamento = ""
	referente_meses = ""
	valor_porcentagem  = ""
	entrada = ""

	doc_path = DocxTemplate(r"./dataset/contrato_renegociacao.docx")
	context = {"nome": nome, "cpf": cpf, "unidade": unidade, "residencial": residencial, "credor" : credor, "cnpj": cnpj,
	           "valor": valor, "data_entrada": data_entrada, "numero_parcelas": numero_parcelas, "valor_parcelas": valor_parcela, "entrada": entrada,
	           "primeiro_pagamento": primeiro_pagamento, "ultimo_pagamento": ultimo_pagamento, "referente_meses": referente_meses, "valor_procentagem": valor_porcentagem}
	doc_path.render(context)
	doc_path.save("teste.docx")
