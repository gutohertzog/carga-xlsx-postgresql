'''
primeira versão do script para realizar a carga
'''
import time
import openpyxl

def le_xlsx(arquivo:str) -> list:
    ''' abre, lê e salva numa lista o conteúdo do arquivo xlsx '''
    planilha = openpyxl.load_workbook(arquivo)
    aba = planilha.active

    matrix:list = []

    for linha in aba.iter_rows(values_only=True):
        matrix.append(list(linha))

    return matrix


if __name__ == '__main__':
    ''' cria as queries para serem executadas no bd a partir do pgAdmin '''
    inicio = time.time()  # cronômetro

    nome_arq = 'arquivo-carga.xlsx'

    matrix = le_xlsx(nome_arq)

    # a partir da matrix, cria um arquivo .sql com todas as queries de
    # atualização dos registros
    with open('atualiza.sql', 'w', encoding='utf-8') as arq:
        for linha in matrix[4:]:
            if linha[6] == 'ATIVO':
                sit = 2
            elif linha[6] == 'INATIVO':
                sit = 4
            elif linha[6] == 'SEM INFORMACAO':
                sit = 5
            else:
                sit = None
                print(f'erro na situação {linha =}')
                exit()
            query = f'update <tabela> set vinculo = {sit} '
            query += f"where cpf = '{linha[4]:0>11}' and curso_id in (5,6);"
            arq.write(query + '\n')

    fim = time.time()
    print(f'terminado em {fim - inicio} segundos')

