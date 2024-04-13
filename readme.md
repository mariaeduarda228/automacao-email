# Primeira aula

```
import pandas as pd
```

## 1 - Importando Dados

```
data = pd.read_excel("data/VendaCarros.xlsx")
print(data)
```

## 2 - Lista os primeiros registros

```
print(data.head())
```

## 3 - Lista os últimos registros

```
print(data.tail())
```

## 4 - Contagem de valores por Fabricante

```
print(data["Fabricante"].value_counts())
```

# Segunda Aula

## 1 - Importando Dados

```
nome = 'Rodrigo'
print(type(nome))
```

```
<class 'str'>
```

```
nome = 'Rodrigo'
print(type(data))
```

```
<class 'pandas.core.frame.DataFrame'>
```

## 2 - Selecionando colunas

```
print(data["Estado"])
```

## 3 - Selecionando colunas específicas do DataFrame

```
df = data[["Fabricante", "ValorVenda", "Ano"]]
```

## 4 - Criando a tabela pivô

```
pivot_table = df.pivot_table(
    index="Ano",
    columns="Fabricante",
    values="ValorVenda",
    aggfunc="sum"
)
```

## 5 - Exportando tabela pivô emarquivo excel

```
pivot_table.to_excel("data/pivot_table.xlsx", "Relatorio")
```

# Terceira Aula

```
from openpyxl import load_workbook
```

## 1 - Lê pasta de trabalho e planilha

```
wb = load_workbook("data/pivot_table.xlsx")
print(wb)
```

```
wb = load_workbook("data/pivot_table.xlsx")
sheet = wb["Relatorio"]
print(sheet)
```

## 2 - Acessando um valor específico

```
print(sheet["A3"].value)
print(sheet["B3"].value)
```

## 3 - Iterando valores por meio de Loop

```
for i in range(2, 5):
    print(i)
```

```
for i in range(2, 6):
    ano = sheet["A%s" % i].value
    am = sheet["B%s" % i].value
    bt = sheet["C%s" % i].value
    print("{0} o Aston Martin vendeu {1} e o Bentley vendeu {2}".format(ano, am, bt))
```

# Quarta Aula

```
from openpyxl import load_workbook
```

## 1-Lê pasta de trabalho e planilha

```
wb = load_workbook("data/pivot_table.xlsx")
sheet = wb["Relatorio"]
```

## 2-Referências das linhas e colunas

```
min_column = wb.active.min_column
max_column = wb.active.max_column
print(min_column)
print(max_column)
```

```
min_column = wb.active.min_column
max_column = wb.active.max_column
min_row = wb.active.min_row
max_row = wb.active.max_row
print(min_row, max_row)
```
