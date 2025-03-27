# sistema_impressao_etiquetas
 Sistema de impressão de etiquetas - Corte

1- Carregar e filtrar o Excel: (feito o controle)

Usar a biblioteca pandas para carregar o arquivo Excel. (Pandas carregado)

Filtrar a lista de ordens pelo "Work Center" do corte. (filtro realizdo E103_CUT)

Replicar o número de série de acordo com a quantidade de confirmações esperadas.(correção do sequencial)

2 - Gerar etiquetas para impressão com QR Code:

Criar um modelo proprio para impressão com o QR code (feito o modelo, preciso aprimorar) (melhorado, porem conseguimos deixar ainda melhor!)

Usar a biblioteca zebra (ou rawprint) para formatar as etiquetas no padrão da impressora Zebra. (feito)

Gerar uma saída que a impressora Zebra aceite diretamente (como ZPL, o comando padrão dela). (feito)

PTVDCP0048 - 192.168.221.81
PTVDCP0050 - 192.168.221.99
PTVDCP0051 - 192.168.221.69
PTVDCP0052 - impressora desligada

http://ptvdcp00XX.enercon.de/server/CFGPAGE.htm

configurações da etiqueta, arquivo zpl:
^XA
^CI28
^FO20,20^A0N,80,60^FDDescrição: Caixa Balanç. BCH E115EP3RB03^FS
^FO20,150^A0N,60,40^FDMaterial: 717803^FS
^FO20,250^A0N,80,60^FDNº Série: EVC0333^FS
^FO750,100^BQN,2,10^FDMA,717803;EVC0333^FS
^XZ


3 - Adicionar uma interface (extra):

Se quiser algo mais prático, podemos montar uma interface simples com Tkinter ou até uma interface web com Flask.
Tela inicial com um campo para o número da ordem.

Botão para imprimir as etiquetas.

Mostra se a ordem foi impressa ou não.