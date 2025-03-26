# sistema_impressao_etiquetas
 Sistema de impressão de etiquetas - Corte

1- Carregar e filtrar o Excel:

Usar a biblioteca pandas para carregar o arquivo Excel.

Filtrar a lista de ordens pelo "Work Center" do corte.

Replicar o número de série de acordo com a quantidade de confirmações esperadas.

2 - Conectar com Google Sheets:

Usar a biblioteca gspread para acessar o Google Sheets.

Pegar o modelo de etiqueta e gerar uma nova aba ou arquivo com os números de série organizados.

3 - Gerar etiquetas para impressão:

Usar a biblioteca zebra (ou rawprint) para formatar as etiquetas no padrão da impressora Zebra.

Gerar uma saída que a impressora Zebra aceite diretamente (como ZPL, o comando padrão dela).

4 - Adicionar uma interface (extra):

Se quiser algo mais prático, podemos montar uma interface simples com Tkinter ou até uma interface web com Flask.