from juntaPlanilha import juntaPlanilha

path = "Path/to/your/excel/files"  # Substitua pelo caminho correto (OBS: use barras normais ou duplas)
linha_inicial = 6   # Linha inicial para leitura (0=primeira linha, 1=segunda linha, etc.)
linha_final = 8   # Linha final para leitura (0=primeira linha, 1=segunda linha, etc.)
# Se linha_final = 0, lê até o final do arquivo
coluna_inicial = 0     # primeira coluna a ser lida (A=0, B=1, C=2, D=3, E=4) :: A 
coluna_final = 5     # Última coluna a ser lida (A=0, B=1, C=2, D=3, E=4) :: E
# Se coluna_final = 0, lê até o final do arquivo
parar_no_vazio = True # True, False. Se True, para na primeira linha vazia encontrada
colunas = ["ID", "NOME", "X", "Y", "Z"] # Defina os nomes das colunas conforme necessário 
nome_arquivo_saida = "planilha_junta.xlsx" # Nome do arquivo de saída


# Chama a função para juntar as planilhas
juntaPlanilha(
    path=path,
    linha_inicial= linha_inicial,
    linha_final= linha_final,
    coluna_inicial= coluna_inicial ,
    coluna_final= coluna_final,
    colunas= colunas,
    nome_arquivo_saida= nome_arquivo_saida,
    parar_no_vazio= parar_no_vazio
)