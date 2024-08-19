'''
v.1.2

é esperado que o arquivo venha com 9 colunas [0..8],
sendo apenas as duas abaixo importantes, por hora.
    - [4] cpf
    - [6] vínculo
'''

import time
import openpyxl
import psycopg2

from getpass import getpass

inicio = time.time()  # cronômetro tempo de execução

ARQ_XLSX = 'arquivo-carga.xlsx'

print('lendo o arquivo .xlsx')
planilha = openpyxl.load_workbook(ARQ_XLSX)
aba = planilha.active

matrix = []
for linha in aba.iter_rows(values_only=True):
    matrix.append(linha)

# contador para os registros atualizados
contador = 0
try:
    print('\nestabelecendo conexão com o banco de dados\n')
    usuario = input('\tdigite o usuário : ')
    senha = getpass('\tdigite a senha : ')
    ip_server = input('\tdigite o IP : ')
    banco = input('\tdigite o banco : ')
    porta = 5432

    conexao = psycopg2.connect(
        database = banco,
        user     = usuario,
        password = senha,
        host     = ip_server,
        port     = porta)

    print('\nconxão bem sucedida')
    cursor = conexao.cursor()

    print('\natualizando registros')
    for linha in matrix:
        # descarta os cabeçalhos
        if isinstance(linha[4], str):
            continue

        if linha[6] == 'ATIVO':
            vinc = 2
        elif linha[6] == 'INATIVO':
            vinc = 4
        elif linha[6] == 'SEM INFORMACAO':
            vinc = 5
        else:
            raise TypeError(linha)

        cpf = str(linha[4]).zfill(11)
        query = 'update <tabela> set vinculo = %s'
        query += 'where cpf = %s and curso_id in (5,6)'

        cursor.execute(query, (vinc, cpf))
        contador += cursor.rowcount

    coexao.commit()
    print('operação finalizada')

    cursor.close()
    conexao.close()
    print('conexão com o banco fechada')

except TypeError as erro:
    print('\n\t\terro de vínculo desconhecido\n')
    print(f'{erro = }\n')

except (Exception, psycopg2.Error) as erro:
    print('\n\terro ao tentar atualizar no banco\n')
    print(f'{erro = }\n')

fim = time.time()
print(f'terminado em {fim - inicio:.2f} segundos')

