import pandas as pd
import os


def juntaPlanilha(path, linha_inicial, linha_final, coluna_inicial, coluna_final, colunas, nome_arquivo_saida, parar_no_vazio=False):
    # Tentativa de 'arrumar' o caminho para pasta
    caminho_pasta = os.path.normpath(path)

    # Lista para armazenar os DataFrames
    dataframes = []

    cabecalho_definido = False
    colunas_cabecalho = None

    # Percorre todos os arquivos Excel na pasta
    for arquivo in os.listdir(caminho_pasta):
        if arquivo.startswith("~$"):  # Ignora arquivos temporários do Excel
            continue
        if arquivo.endswith(".xls") or arquivo.endswith(".xlsx"):
            caminho_arquivo = os.path.join(caminho_pasta, arquivo)

            # Lê todas as planilhas do arquivo
            xls = pd.ExcelFile(caminho_arquivo)
            for nome_planilha in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=nome_planilha, header=None)  # Lê sem cabeçalhos

                # Seleciona a partir da linha passada e pega colunas (0 a 4)
                df_dados = df.iloc[int(linha_inicial)-1:(None if linha_final == 0 else int(linha_final)), int(coluna_inicial):int(coluna_final)]

                # Encontra a primeira linha totalmente vazia
                primeira_linha_vazia = df_dados[df_dados.isnull().all(axis=1)].index
                if not primeira_linha_vazia.empty and parar_no_vazio:    # Se parar no vazio for True
                    df_dados = df_dados.loc[:primeira_linha_vazia[0] - 1]  # Mantém só até antes da linha vazia

                # Encontra a primeira linha que contém o cabeçalho válido e remove as anteriores
                primeira_linha_valida = df_dados[df_dados.iloc[:, 0] == colunas[0]].index
                if not primeira_linha_valida.empty:
                    df_dados = df_dados.loc[primeira_linha_valida[0] + 1 :]  # Remove o cabeçalho duplicado


                # Define o cabeçalho apenas uma vez (pegando do primeiro arquivo processado)
                if not cabecalho_definido:
                    colunas_cabecalho = colunas
                    cabecalho_definido = True

                df_dados.columns = colunas_cabecalho

                # Adiciona colunas extras para identificação
                #df_dados["Arquivo_Origem"] = arquivo
                #df_dados["Planilha_Origem"] = nome_planilha

                dataframes.append(df_dados)

    # Junta todos os DataFrames
    df_final = pd.concat(dataframes, ignore_index=True)

    # Salva em um novo Excel
    df_final.to_excel(f"{nome_arquivo_saida}.xlsx", index=False)

    print("✅ Arquivos unidos com sucesso, sem informações extras!")
