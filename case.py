'''
Utilizei de um ambiente virtual para a instalação de bibliotecas externas.
os.walk retorna um elemento gerador, ao envolver a função com o next eu estou desempacotando a tupla do gerador em 3 variáveis para utilizar no código.
Estou utilizando de list comprehension para criar uma lista de arquivos que atendem aos critérios solicitados.
Crio o arquivo excel atráves do pandas por estar mais habituado pela utilização na rotina do trabalho, mas não encontrei na biblioteca uma forma de adicionar colunas ao final, como solicitado. Para não utilizar a biblioteca openpyxl a opção seria reescrever o dataframe n vezes ou somente atualiza-lo ao final das iterações.
Ao invés de abrir os arquivos e salvar como, optei por renomea-los, uma vez que o output é o mesmo e polpa tempo de processamento.
Finalizei encerrando o arquivo de excel que estava pré carregado.
'''
import os
import pandas as pd
from openpyxl import load_workbook

#Verificar arquivos na pasta
path, dirs, files = next(os.walk('.\RPA-Artigo'))
files_criterio = [file for file in files if file[0].isdigit() & file.endswith('.pdf')]

#Criar dataframe e arquivo excel
data = {'Nome do documento':[], 'Status':[]}
df = pd.DataFrame(data)
path_excel = os.path.join(path, 'Relatório de execução.xlsx')
df.to_excel(path_excel, index= False)

#Carregar arquivo excel
workbook = load_workbook(filename = path_excel)
sheet = workbook['Sheet1']

for file in files_criterio:
    try:
        #Renomear
        old_name = os.path.join(path, file)
        file_number_index = file.find('_')
        rule = file[:file_number_index] + '_Página 7 - Modificado.pdf'
        new_name = os.path.join(path, rule)
        os.rename(old_name, new_name)
        message = 'documento alterado'
    except Exception as e:
        message = str(e)

    #Inserir no excel
    row = [file, message]
    sheet.append(row)

workbook.save(path_excel)
workbook.close()