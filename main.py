'''
v.1.1

é esperado que o arquivo venha com 9 colunas [0..8],
sendo apenas as duas abaixo importantes, por hora.
    - [4] cpf
    - [6] situação
'''

import time
import openpyxl

inicio = time.time()  # cronômetro tempo de execução

ARQ_SQL = 'atualiza.sql'
ARQ_XLSX = 'arquivo-carga.xlsx'

planilha = openpyxl.load_workbook(ARQ_XLSX)
aba = planilha.active

matrix = []
for linha in aba.iter_rows(values_only=True):
    matrix.append(linha)

# a partir da matrix, cria todas as queries
queries = []
for linha in matrix:
    # descara os cabeçalhos
    if len(linha < 8):
        continue
    if len(linha[4]) <= 3:
        continue

    if linha[6] == 'ATIVO':
        sit = 2
    elif linha[6] == 'INATIVO':
        sit = 4
    elif linha[6] == 'SEM INFORMACAO':
        sit = 5
    else:
        print('\n\n\t\terro de vínculo desconhecido\n')
        print(f'{linha = }\n')
        exit()

    cpf = str(linha[4]).zfill(11)
    query = f'update <tabela> set vinculo = {sit} '
    query += f"where cpf = '{cpf}' and curso_id in (5,6);"

with open(ARQ_SQL, 'w', encoding='utf-8') as arq:
    for query in queries:
        arq.write(query + '\n')

fim = time.time()
print(f'terminado em {fim - inicio:.2f} segundos')

