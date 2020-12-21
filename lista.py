print("================================================================================")
print("=             PYTHON v3.9                                                      =")
print("=             GRUPO B&K - MALA DIRETA                                          =")
print("=             v1.0                                                             =")
print("=             Github: https://github.com/masterfake2/mala-direta               =")
print("================================================================================")

import os
import pandas as pd
from docxtpl import DocxTemplate
import docx2pdf
from datetime import datetime
import cloudmersive_convert_api_client
from cloudmersive_convert_api_client.rest import ApiException
from pprint import pprint
#Modelos - EXECUÇÃO DE TÍTULO EXECUTIVO DE COTAS CONDOMINIAIS
#Modelos - CONTRATO DE RENEGOCIAÇÃO DE DÍVIDA


##CODIGO DE ERROS 
## 4 - ERRO GERAL VOLTA PARA O MENU PRINCIPAL

#Classe PAI responsável por pegar os dados
class Dados:
    #Dados devedor
    devedor_nome = ""
    devedor_cpf  = ""

    #Dados Sindico
    sindico_nome = ""
    sindico_cpf  = ""

    #Dados residencial
    residencial_cnpj = ""
    residencial_unidade = ""
    residencial_residencial = ""

    def __init__(self, cpf):
        self.cpf = cpf

    def buscarDataSet(self):
        try:
            excel_path = r"./dataset/base.xls"
            self.dataset = pd.read_excel(excel_path)
        except Exception as e:
            return 4

    def procurarInformacao(self):
        try:
	        # Faz a busca no dataset e retorna a linha
            self.row = self.dataset.loc[self.dataset["CPF DEVEDOR"] == self.cpf]
        except Exception as e:
            print(e)
            # Linha não foi encontrada ou seja registro não existe, ou chave não existe
            self.row = False
            return 4
        if(len(self.row) < 1):
            print("[SISTEMA] CPF NÃO ENCONTRADO!")
            return 4
            
        print('---------------------------')
        texto = "Encontrado" if len(self.row) <= 1 else "Encontrados"
        print("---- {} {} ---------".format(len(self.row), texto))
        print('---------------------------')
        print("Nome: {}".format(self.row.get("NOME").values[0]))    
        print("---------------------------")

    def setDados(self):
        #Devedor
        self.devedor_nome = self.row.get("NOME").values[0]
        self.devedor_cpf = self.row.get("CPF DEVEDOR").values[0]

        #Residencial
        self.residencial_cnpj = self.row.get("CNPJ").values[0]
        self.residencial_unidade = self.row.get("UNIDADE").values[0]
        self.residencial_residencial = self.row.get("RESIDENCIAL").values[0]

        #Sindico
        self.sindico_nome = self.row.get("SINDICO")
        self.sindico_cpf  = self.row.get("CPF SINDICO")

class Arquivos:
    diretorio_default = r"./mala-direta-arquivos/"
    agora = datetime.now()
    ano   = agora.year
    mes   = agora.month
    dia   = agora.day

    if(mes == 1):
        mes = "Janeiro"
    elif(mes == 2):
        mes = "Fevereiro"
    elif(mes == 3):
        mes = "Março"
    elif(mes == 4):
        mes = "Abril"
    elif(mes == 5):
        mes = "Maio"
    elif(mes == 6):
        mes = "Junho"
    elif(mes == 7):
        mes = "Julho"
    elif(mes == 8):
        mes = "Agosto"
    elif(mes == 9):
        mes = "Setembro"
    elif(mes == 10):
        mes = "Outubro"
    elif(mes == 11):
        mes = "Novembro"
    elif(mes == 12):
        mes = "Dezembro"
    
    caminho = ""
    def criaDiretorio(self, nome):
        try:
            os.mkdir(self.diretorio_default + nome)
        except:
            pass
        #Se criou o dir principal então cria os próximos
        try:
            os.mkdir(self.diretorio_default + nome + "/CONTRATOS")
        except:
            pass
        
        try:
            os.mkdir(self.diretorio_default + nome + "/EXECUÇÕES")
        except:
            pass
        print("[SISTEMA] Diretórios criados!")
        print("[SISTEMA] Continuando script.")
        print("[SISTEMA] Concluído com sucesso!")

    def criarModeloContratoDeRenegociacao(self, dadosPlanilha, dadosInputs):
        contexto = {"devedor_nome": dadosPlanilha.devedor_nome, "devedor_cpf": dadosPlanilha.devedor_cpf, "residencial_unidade": dadosPlanilha.residencial_unidade,
                    "residencial_residencial": dadosPlanilha.residencial_residencial, "residencial_cnpj": dadosPlanilha.residencial_cnpj, "divida_montante_total": dadosInputs.divida_montante_total,
                    "divida_referente_meses": dadosInputs.divida_referente_meses, "divida_entrada": dadosInputs.divida_entrada, "divida_data_pagamento_da_entrada": dadosInputs.divida_data_pagamento_da_entrada,
                    "divida_numero_de_parcelas": dadosInputs.divida_numero_de_parcelas, "divida_valor_parcelas": dadosInputs.divida_valor_parcelas, "divida_data_primeiro_pagamento_parcelas": dadosInputs.divida_data_primeiro_pagamento_parcelas,
                    "divida_data_ultimo_pagamento_parcelas": dadosInputs.divida_data_ultimo_pagamento_parcelas, "dia": self.dia, "mes": self.mes, "ano": self.ano
                    }
        doc_path = DocxTemplate(r"./dataset/contrato-modelo-renegociacao.docx")
        doc_path.render(contexto)
        self.caminho = self.diretorio_default + dadosPlanilha.devedor_nome + "/CONTRATOS/" + dadosPlanilha.devedor_nome + ".docx"
        self.caminho_pdf = self.diretorio_default + dadosPlanilha.devedor_nome + "/CONTRATOS/" 
        doc_path.save(self.caminho)
        print("[SISTEMA] Caminho para o arquivo: {}".format(self.caminho))
        configuration = cloudmersive_convert_api_client.Configuration()
        configuration.api_key['Apikey'] = '5e1c7861-12ef-43ae-a5d7-6d21c0c52c3a'
        api_instance = cloudmersive_convert_api_client.ConvertDocumentApi(cloudmersive_convert_api_client.ApiClient(configuration))
        input_file = self.caminho
        api_response = api_instance.convert_document_docx_to_pdf(input_file)
        #Salvar como PDF
        pdf = open(self.caminho_pdf + dadosPlanilha.devedor_nome + ".pdf", "wb")
        pdf.write(api_response)
        pdf.close()
        
class Modelos:
    divida_montante_total = ""
    divida_referente_meses = ""
    divida_entrada = ""
    divida_data_pagamento_da_entrada = ""
    divida_numero_de_parcelas = ""
    divida_data_primeiro_pagamento_parcelas = ""
    divida_data_ultimo_pagamento_parcelas = ""
    divida_valor_parcelas = ""
    #Função responsável pelo contrato de renegociacao
    def setContratoDeRenegociacao(self):
        #Pegar dados específicos do contrato de renegociação
        valida = 0
        try:
            print("> Digite o valor total do montante devido, no formato de exemplo: R$ 235,96")
            self.divida_montante_total = str(input("> R$ "))
        except:
            return 4
        print("\n")
        try:
            print("> Digite a linha de referencia mensal, no formato de exemplo: 08 e 10/2020 e acordo dos meses 02, 03, 04, 05 e 08/2020")
            self.divida_referente_meses = str(input("> Referente aos meses : "))
        except:
            return 4
        print("\n")
        try:
            print("> Digite o valor das parcelas, no formato de exemplo: R$ 79,16")
            self.divida_valor_parcelas = str(input("> R$ "))
        except:
            return 4
        print("\n")
        try:
            print("> Digite o valor da entrada, no formato de exemplo: R$ 79,16")
            self.divida_entrada = str(input("> R$ "))
        except:
            return 4
        print("\n")
        try:
            print("> Data do pagamento da entrada, no formato de exemplo: 07/12/2020")
            self.divida_data_pagamento_da_entrada = str(input(">  "))
        except:
            return 4
        print("\n")
        try:
            print("> Número de parcelas, no formato de exemplo: 2")
            self.divida_numero_de_parcelas = str(input(">  "))
        except:
            return 4
        print("\n")
        try:
            print("> Data do primeiro pagamento das parcelas, no formato de exemplo: 07/01/2021")
            self.divida_data_primeiro_pagamento_parcelas = str(input(">  "))
        except:
            return 4
        print("\n")
        try:
            print("> Data do último pagamento das parcelas, no formato de exemplo: 07/02/2021")
            self.divida_data_ultimo_pagamento_parcelas = str(input(">  "))
        except:
            return 4


    #Função responsável pelo contrato de Execusão de titulos
    def setExecusaoDeTituloExecutivo():
        pass

class Menu():
    def showMenu(self):
        print("\n1 - Contrato de renegociacao")
        print("2 - Execusao de titulos *Desativado esperando validação")
        print("3 - Consultar dados *Desativado esperando validação")

    def showErro(self):
        print("[SISTEMA] Erro, favor verificar valor digitado!\n")

    def setCpf(self):
        cpf = str(input("[SISTEMA] CPF para buscar > "))
        return cpf

    def verificarOpcao(self, opcao):
        if(opcao == 1): #Contrato de renegociacao
            dadosPlanilha = Dados(self.setCpf())
            dadosInputs   = Modelos()
            arquivos      = Arquivos() #Cria a instância da classe arquivos, resposável por toda a transição dos dados para às informações do word -> pdf
            if(dadosPlanilha.buscarDataSet() != 4): #Verifica se conseguiu encontrar o arquivo da fonte de dados
                if(dadosPlanilha.procurarInformacao() != 4): #Pega informacao do CPF digitado 
                    dadosPlanilha.setDados() #Seta as variaveis com os dados encontrados relacionados ao devedor, residencial e ao sindico
                    dadosInputs.setContratoDeRenegociacao() #Seta as variaveis relacionadas a divida
                    arquivos.criaDiretorio(dadosPlanilha.devedor_nome) #Cria o diretorio principal do devedor selecionado pelo cpf
                    arquivos.criarModeloContratoDeRenegociacao(dadosPlanilha, dadosInputs)
                else:
                    pass #volta para o menu principal
            else: 
                pass #volta para o menu principal
        else:
            self.showErro()

#Cria menu 
menu = Menu()
while True:
    menu.showMenu()
    opcao = int(input("Escolha > "))
    menu.verificarOpcao(opcao)