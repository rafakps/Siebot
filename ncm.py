import datetime
import os

import openpyxl as opxl
import requests
import pandas as pd
import json



class AliquotaNCM:
    """
    Funções:

    verificarTabela - Verifica a data de hoje e já passa o nome da tabela de hoje

    baixarTabela - Baixa uma tabela com as alíquotas de ICMS e salva com a data de hoje

    aliquotaNCM - Retorna o valor da Aliquota de IPI de um determinado produto

    descricacaoNCM - Retorna a Descrição de um NCM

    """

    def verificarTabela():
        #Verifica a data de hoje e já passa o nome da tabela de hoje
        hoje = datetime.date.today().strftime('%d-%m-%Y')
        nome_tabela = 'Tabela Diaria\\TabelaNCM%s.xlsx' % (hoje)
        return nome_tabela


    def baixarTabela():
        #Verifica se já há uma tabela baixada para ser utilizada hoje
        nome_tabela = AliquotaNCM.verificarTabela()

        if os.path.exists(nome_tabela):
            return print("Tabela já exite")
        else:
            # Função baixa a tabela do site do gov e tira os cabeçalhos
            tabelaNCM = requests.get('https://www.gov.br/receitafederal/pt-br/acesso-a-informacao/legislacao/documentos-e-arquivos/tipi-em-excel.xlsx')
            if tabelaNCM.status_code == requests.codes.OK:
                with open('tabelaNCM.xlsx', 'wb') as novaTabela:
                    novaTabela.write(tabelaNCM.content)
                    #carrega o arquivo
                    tabelaDados1 = opxl.load_workbook(filename='tabelaNCM.xlsx')
                    tabelaDados = tabelaDados1.worksheets[0]
                    #Deleta o cabeçalho da tabela, para ficar apenas com os dados
                    tabelaDados.delete_rows(1,7)
                    #Atrui zero as células de Alíquota vazias
                   # for i in range(1,tabelaDados.max_row):
                    #    if tabelaDados['D'][i].value == '':
                     #       tabelaDados['D'][i].value = 0


                    """Dados da tabela
                       colunas: NCM (1) | EX (2) | Descrição (3) | Aliquota(%) (4)
                    """
                    #Salva uma nova tabela com a data de hoje e cria um json

                    df = pd.DataFrame(tabelaDados.values)

                    for i in range(df.shape[0]):
                        if df[3][i] == '':
                            df[3][i] = 0



                    hoje = datetime.date.today().strftime('%d-%m-%Y')
                    nomeJson = 'Tabela Diaria\\TabelaNCM%s.json' % (hoje)
                    df.to_json(nomeJson,orient='columns')

                    return tabelaDados1.save(nome_tabela)

            else:
                tabelaNCM.raise_for_status()

    def aliquotaNCM(codigoNCM):
        #verifica a aliquota contida no json
        hoje = datetime.date.today().strftime('%d-%m-%Y')
        nome_json = 'Tabela Diaria\\TabelaNCM%s.json' % (hoje)
        # verifica se o json de hoje já está baixado caso contrário ele faz um novo download
        if os.path.exists(nome_json):
            df = pd.read_json(nome_json)

            #codigoNCM1 = AliquotaNCM.entradaDados(codigoNCM) -- irei implmentar
            for i in range(df.shape[0]):
                if codigoNCM == df[0][i]:
                    return df[3][i]
                    """
                    if df[3][i] == None:# or df[3][i] == '':
                        return 0
                    else:
                        return df[3][i]
                    """
        else:
            AliquotaNCM.baixarTabela()
            return AliquotaNCM.aliquotaNCM(codigoNCM)

        return 'Código não encontrado'


    def descricaoNCM(codigoNCM):
        #Verifica a descricao do código NCM
        hoje = datetime.date.today().strftime('%d-%m-%Y')
        nome_json = 'Tabela Diaria\\TabelaNCM%s.json' % (hoje)
        # verifica se o json de hoje já está baixado caso contrário ele faz um novo download
        if os.path.exists(nome_json):
            df = pd.read_json(nome_json)
            #codigoNCM1 = AliquotaNCM.entradaDados(codigoNCM) -- irei implementar
            for i in range(df.shape[0]):
                if codigoNCM == df[0][i]:
                    return df[2][i]
        else:
            AliquotaNCM.baixarTabela()
            return AliquotaNCM.aliquotaNCM(codigoNCM)

        return 'Código não encontrado'

    def aliquotaNCM2(codigoNCM2):
        #Verifica a descricao do código NCM
        hoje = datetime.date.today().strftime('%d-%m-%Y')
        nome_json = 'Tabela Diaria\\TabelaNCM%s.json' % (hoje)
        # verifica se o json de hoje já está baixado caso contrário ele faz um novo download
        if os.path.exists(nome_json):
            df = pd.read_json(nome_json)
            for i in range(df.shape[0]):
                if df[3][i] == '' or df[3][i] == df[3][1]:
                    df[3][i] = 0


            #codigoNCM1 = AliquotaNCM.entradaDados(codigoNCM) -- irei implementar
            for i in range(df.shape[0]):
                if codigoNCM2 == df[0][i]:
                    return df.iloc[3][i]
        else:
            AliquotaNCM.baixarTabela()
            return AliquotaNCM.aliquotaNCM(codigoNCM2)

        return 'Código não encontrado'

    def entradaDados(texto):
        entrada = []
        entrada1 = []
        ab = 0

        # capta somente os números para depois transformalos no formato de entrada da tabela
        # irei implementar de maneira melhor
        for i in range(len(texto)):
            if texto[i] != '.':
                entrada.append(texto[i])
                if len(entrada) == 8:
                    entrada.insert(4,'.')
                    entrada.insert(7,'.')
                    return "".join(entrada)



#testes que fiz para testar a classe
"""
lista = ['1.01','1.02','1.03','1.04','1.05','1.06','2.01','2.02','2.03','2.04']
for i in range(10):
    print(AliquotaNCM.descricaoNCM(lista[i]), AliquotaNCM.aliquotaNCM(lista[i]))

df = pd.read_json('Tabela Diaria\\TabelaNCM02-08-2022.json')

print(df)
print(df[0][1],df[1][1],df[2][1],df[3][1])
print('ok')

for i in range(df.shape[0]):
#for i in range(20):
    if df[3][i] == df[3][1] or df[3][i] == '':
        df[3][i] = 0

print(df)

for i in range(100):
    print(AliquotaNCM.descricaoNCM(df[0][i]), AliquotaNCM.aliquotaNCM(df[0][i]))
"""
"""
print(AliquotaNCM.entradaDados('0202.10.00.123456788910'))

rafael = AliquotaNCM.entradaDados('0202.10.00.123456788910')

print(AliquotaNCM.descricaoNCM('0202.10.00'),AliquotaNCM.aliquotaNCM('0202.10.00'))


#- Carcaças e meias-carcaças

"""



