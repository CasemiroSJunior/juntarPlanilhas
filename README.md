# 📊 Junta Planilhas Excel com Python

Este script permite **unir diversas planilhas Excel de uma pasta** em um único arquivo `.xlsx`, com controle de quais linhas e colunas serão lidas, cabeçalhos personalizados e possibilidade de parar na primeira linha vazia.

---

## ✅ Funcionalidades

- Junta arquivos `.xlsx` e `.xls` contidos em uma pasta
- Lê várias planilhas dentro de cada arquivo Excel
- Permite configurar o intervalo de linhas e colunas a serem lidas
- Define nomes personalizados para as colunas
- Opcional: parar na primeira linha vazia encontrada
- Gera um arquivo Excel final com todos os dados unidos

---

## ⚙️ Requisitos

Instale os pacotes necessários com:

```bash
pip install -r requirements.txt
```

**requirements.txt:**

```
et_xmlfile==2.0.0
numpy==2.2.4
openpyxl==3.1.5
pandas==2.2.3
python-dateutil==2.9.0.post0
pytz==2025.2
six==1.17.0
tzdata==2025.2
```

---

## 🧠 Como usar

### 1. Ajuste o caminho e as configurações no seu `main.py`:

```python
from juntaPlanilha import juntaPlanilha

path = "Path/to/your/excel/files"  # <- Substitua pelo caminho correto da pasta
linha_inicial = 6
linha_final = 8
coluna_inicial = 0
coluna_final = 5
parar_no_vazio = True
colunas = ["ID", "NOME", "X", "Y", "Z"]
nome_arquivo_saida = "planilha_junta.xlsx"

juntaPlanilha(
    path=path,
    linha_inicial=linha_inicial,
    linha_final=linha_final,
    coluna_inicial=coluna_inicial,
    coluna_final=coluna_final,
    colunas=colunas,
    nome_arquivo_saida=nome_arquivo_saida,
    parar_no_vazio=parar_no_vazio
)
```

> ⚠️ **Dica:** use **barras normais (`/`) ou duplas (`\\`)** no caminho da pasta para evitar erros de leitura.

### 2. Execute:

```bash
python main.py
```

O resultado será salvo como `planilha_junta.xlsx` no mesmo diretório do script.

---

## 📁 Estrutura do Projeto

```
📁 seu-projeto/
├── juntaPlanilha.py
├── main.py
├── requirements.txt
└── README.md
```

---

## 📌 Observações

- A função ignora arquivos temporários do Excel (que começam com `~$`)
- Lê todas as planilhas dentro de cada arquivo Excel da pasta
- O cabeçalho é aplicado com base nos nomes definidos em `colunas`
- O script é compatível com arquivos `.xls` e `.xlsx`

---

## 📃 Licença

Este projeto é de uso livre para fins educacionais e pessoais. Sinta-se à vontade para contribuir ou adaptar conforme sua necessidade.

---
