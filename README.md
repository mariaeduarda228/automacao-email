# Automação de Planilhas com o Python

 A XYZ Corporation International atua como uma revendedora de automóveis de luxo com sede em São Paulo. A empresa começou sua operação no Brasil em 2012 e precisa apresentar os resultados da equipe comercial para o novo CEO da empresa.
 
 O CEO deseja saber como foram as vendas dos carros por cada fabricante em cada um dos anos. Ele precisa que essa informação seja apresentada em forma de gráficos e precisa ter também o total geral de vendas de cada fabricante. O resultado final deve estar numa nova planilha e ser enviada por e-mail para o CEO.

Sua fonte de dados é um arquivo Excel com dados coletados do sistema de vendas e CRM da empresa, com a as seguintes colunas:

| Coluna | Descrição |
|----------|---------|
| DataNotaFiscal | Data de emissão da nota fiscal |
| Fabricante | Fabricante do veículo |
| Estado | Estado onde foi realizada a venda |
| PrecoVenda | Peço de venda do veículo |
| PrecoCusto | Preço de custo do veículo para a empresa |
| TotalDesconto | Total de Desconto fornecido sobre o preço de venda |
| CustoEntrega | Custo de entrega do veículo ao proprietário |
| CustoMaoDeobra | Custo de Mão de Obra (secretária, mecânico, etc...) |
| NomeCliente	| Nome do cliente que comprou o veículo |
| Modelo | Modelo do veículo |
| Ano	| Ano de fabricação do veículo |

Com o objetivo de fornecer as informações para o CEO, a automação deve seguir alguns passos:
1.	Deve ser gerado uma tabela pivô em Python apenas com as colunas: Fabricante, PrecoVenda e Ano. A coluna Ano deve ser o índice para a nova tabela.
2.	Importar a nova planilha gerada com as novas informações e criar um gráfico de barras para apresentar o total de Vendas por Fabricantes ao longo dos anos.
3.	Adicionar fórmula para totalizar as vendas por cada fabricante.
4.	Enviar por e-mail a planilha que foi automatizada.

