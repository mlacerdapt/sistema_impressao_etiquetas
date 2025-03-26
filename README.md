# sistema_impressao_etiquetas
 Sistema de impressão de etiquetas - Corte

1- Carregar e filtrar o Excel: (feito o controle)

Usar a biblioteca pandas para carregar o arquivo Excel. (Pandas carregado)

Filtrar a lista de ordens pelo "Work Center" do corte. (filtro realizdo E103_CUT)

Replicar o número de série de acordo com a quantidade de confirmações esperadas.(correção do sequencial)

2 - Gerar etiquetas para impressão com QR Code:

Criar um modelo proprio para impressão com o QR code (feito o modelo, preciso aprimorar)

Usar a biblioteca zebra (ou rawprint) para formatar as etiquetas no padrão da impressora Zebra. (feito)

Gerar uma saída que a impressora Zebra aceite diretamente (como ZPL, o comando padrão dela). (feito)

PTVDCP0048 - 192.168.221.81
PTVDCP0050 - 192.168.221.99
PTVDCP0051 - 192.168.221.69
PTVDCP0052 - impressora desligada

http://ptvdcp00XX.enercon.de/server/CFGPAGE.htm


3 - Adicionar uma interface (extra):

Se quiser algo mais prático, podemos montar uma interface simples com Tkinter ou até uma interface web com Flask.