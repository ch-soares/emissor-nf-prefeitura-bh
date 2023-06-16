import gspread
import os
import credenciais


def google_sheets(aba):
    gc = gspread.service_account(filename=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'service_account.json'))
    planilha = gc.open_by_key(credenciais.ID_PLANILHA)
    return planilha.worksheet(aba)


def emite_nf():
    aba = google_sheets(credenciais.ABA_DADOS_NF)
    # Obter todos os valores da planilha
    todos_os_valores = aba.get_all_values()
    valores = [linha for linha in todos_os_valores
               if "Liberado" in linha and "Emitida" not in linha
               ]
    # Obter os cabeçalhos das colunas
    cabecalhos = todos_os_valores[0]

    # Inicializar lista vazia para armazenar os dicionários
    dados = []

    # Percorrer as linhas (excluindo o cabeçalho)
    for linha in valores:
        # Inicializar um dicionário para cada linha
        linha_dict = {}
        # Percorrer as colunas e adicionar os valores ao dicionário
        for indice, valor in enumerate(linha):
            coluna = cabecalhos[indice]
            linha_dict[coluna] = valor
        # Adicionar o dicionário à lista de dados
        dados.append(linha_dict)
        
    return dados


def salvar_nfs_emitidas(nfs_emitidas):
    with open('nfs_emitidas.txt', 'w') as arquivo:
        for nf in nfs_emitidas:
            linha = ','.join(str(elemento) for elemento in nf)
            arquivo.write(linha + '\n')


def atualizar_planilha():
    aba = google_sheets(credenciais.ABA_NFS_EMITIDAS)
    with open('nfs_emitidas.txt', 'r') as arquivo:
        dados = [linha.replace('\n', '').split(',') for linha in arquivo]
    aba.append_rows(dados)
