import csv
import glob
import os
import sys
from datetime import datetime as dt
from datetime import timedelta
from timeit import default_timer as timer

import xlrd
from tqdm import tqdm

import unicodedata


# Carregar dados da planilha Excell para arquivo CSV

# Execução por linha de comando:
# ------------------------------------------------------------------------------
# 1 - Path Python -D:\apps\bi\python\extract-data\venv\Scripts\python.exe
# 2 - NomeArquivoXLS - Exemplo: 17.12.2021 Base MP Credity.xlsx
# 3 - NomeArquivoCSV - Exemplo: MP Credity.csv
# Localização do Arquivo Definidos:
# 4 - Diretorio de XLS - D:\apps\bi\data\FILES\output
# 5 - Diretorio de CSV - D:\apps\bi\data\FILES\input
# ------------------------------------------------------------------------------
# CMD Line:
# D:\apps\bi\python\extract-data\venv\Scripts\python.exe "D:\apps\bi\python\extract-data\xls2csv.py" "17.12.2021 Base MP Credity.xlsx" "MP Credity.csv"
# ------------------------------------------------------------------------------

def convert_to_non_accent(string):
    """ Function to convert accent characters to non accent
    characters.
    :param string: String to be converted.
    :type string: str
    :return: str
    """
    return ''.join(ch for ch in unicodedata.normalize('NFKD', string)
                   if not unicodedata.combining(ch))


def new_file_name(file_name, seq):
    buffer = file_name[:-4]
    new_seq = '{0:04d}'.format(seq)
    buffer = buffer + "_" + new_seq + ".csv"
    return buffer


def mask_delete_name(file_name):
    buffer = file_name[:-4]
    buffer = buffer + "*" + ".csv"
    return buffer


print("Sistema de BI & Analytics - Kitei");
print("Desenvolvido por OG1 Tecnologia Inf Ltda");
print("---------------------------------------------------------------");
print("Conversão de Arquivo XLSX");
print("---------------------------------------------------------------");
print('\tParâmetros Informados:', str(sys.argv))
print('\tQtde de Parametros:', len(sys.argv), 'arguments.')
if len(sys.argv) < 3:
    print('**** Erro nos parametros informados. Sistema finalizado. ****')
    exit(-1)

print('\tArgumento 1 - XLS:', sys.argv[1])
print('\tArgumento 2 - CSV:', sys.argv[2])

# preparar arquivos de processamento
seq = 1
file_path = 'D:/apps/bi/data/FILES'
file_xls = file_path + '/output/' + sys.argv[1]
file_csv = file_path + '/input/' + sys.argv[2]
new_file_csv = new_file_name(file_csv, seq)

# Limpar arquivos CSV existentes
del_file = 0
path_csv_delete = file_path + '/input/'
delete_file = mask_delete_name(sys.argv[2])
filelist = glob.glob(os.path.join(path_csv_delete, delete_file))
for f in filelist:
    os.remove(f)
    del_file += 1

print("---------------------------------------------------------------");
print('\tExclusão de %3d Arquivo(s) no Destino=%s' % (del_file, delete_file));
print("---------------------------------------------------------------");

# Open the Workbook
print("---------------------------------------------------------------");
print('\tArquivo de Origem =' + file_xls)
print('\tArquivo de Destino=' + file_csv)
print("---------------------------------------------------------------");

now = dt.now()
start = timer()
print('1 - Abrindo arquivo Excel - ', now)
workbook = xlrd.open_workbook(file_xls)
now = dt.now()
end = timer()
print('2 - Arquivo carregado - ', now)

print('3 - Tempo de carregamento do arquivo em minutos: ', timedelta(minutes=end - start))

# Open the worksheet
worksheet = workbook.sheet_by_name('Analitico')
max_rows = worksheet.nrows - 1
print("4 - Qtd Linhas=", max_rows)

# create the csv writer
if os.path.exists(new_file_csv):
    os.remove(new_file_csv)
    print(f"The file: {new_file_csv} is deleted!")

f_csv = open(file=new_file_csv, mode='w', newline='\n', encoding='utf-8')
writer = csv.writer(f_csv, delimiter=";")

# control variables
buffer = ""
cur_row = 1
start_col = 1
end_col = 42
header = []
data = []

now = dt.now()
print('5- Inicio de gravação do arquivo - ', now)

# Ler Header
header_row = []
get_header_row = 0
for cur_col in range(start_col, end_col):
    buffer = worksheet.cell_value(get_header_row, cur_col)
    buffer = convert_to_non_accent(buffer)
    header_row.append(buffer)

# write a row to the csv file
writer.writerow(header_row)

print("Row Header=>", get_header_row)
print("Header=>", header_row)

# Cedente
cur_row += 1
name_cedente = worksheet.cell_value(cur_row, start_col)
print('\t\tColuna Cedente =' + name_cedente)

if not name_cedente:
    print('*** nao há dados para processamento *** ')
    exit(-1)

# Ler Dados

qtd_buffer = 0
for cur_row in tqdm(range(5, max_rows), desc='linhas processadas: '):
    row = []
    # Processar todas as colunas
    for cur_col in range(start_col, end_col):
        buffer = worksheet.cell_value(cur_row, cur_col)
        # Gravar nova linha no arquivo CSV
        row.append(buffer)

    # incluir linha no arquivo
    writer.writerow(row)
    qtd_buffer += 1

    # verificar se já tem 20K registros processados
    if qtd_buffer >= 20000:
        qtd_buffer = 0
        f_csv.close()
        seq += 1
        new_file_csv = new_file_name(file_csv, seq)
        f_csv = open(file=new_file_csv, mode='w', newline='\n', encoding='utf-8')
        writer = csv.writer(f_csv, delimiter=";")
        writer.writerow(header_row)
        # print('\t\tCriado um novo arquivo:', new_file_csv )

print('6 - Final de processamento - total de registros lidos', cur_row)
f_csv.close()
