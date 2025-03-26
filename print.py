import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import qrcode
import socket

# Função para gerar o QR Code
def gerar_qr_code(conteudo):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(conteudo)
    qr.make(fit=True)
    return qr.make_image(fill='black', back_color='white')

# Função para gerar a etiqueta visual
def gerar_etiqueta(descricao, material, numero_serie):
    largura, altura = 1040, 400
    etiqueta = Image.new('RGB', (largura, altura), 'white')
    draw = ImageDraw.Draw(etiqueta)

    try:
        font_descricao = ImageFont.truetype("arial.ttf", 50)
        font_material = ImageFont.truetype("arial.ttf", 40)
        font_numero = ImageFont.truetype("arial.ttf", 60)
    except IOError:
        font_descricao = font_material = font_numero = ImageFont.load_default()

    draw.text((20, 20), f"DESCRIÇÃO: {descricao}", font=font_descricao, fill='black')
    draw.text((20, 150), f"MATERIAL: {material}", font=font_material, fill='black')
    draw.text((20, 300), f"NÚMERO DE SÉRIE: {numero_serie}", font=font_numero, fill='black')

    qr_img = gerar_qr_code(f"{material};{numero_serie}")
    etiqueta.paste(qr_img, (800, 100))
    etiqueta.save(f"etiqueta_{numero_serie}.png")

# Função para gerar o ZPL
def gerar_zpl(descricao, material, numero_serie):
    return (f"^XA\n"
            f"^FO20,20^A0N,50,50^FDDescrição: {descricao}^FS\n"
            f"^FO20,150^A0N,40,40^FDMaterial: {material}^FS\n"
            f"^FO20,300^A0N,60,60^FDNº Série: {numero_serie}^FS\n"
            f"^FO800,100^BQN,2,10^FDMA,{material};{numero_serie}^FS\n"
            f"^XZ\n")

# Função para enviar para a impressora
def imprimir_zebra(ip, porta, zpl_code):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.connect((ip, porta))
        sock.sendall(zpl_code.encode())

# Carregar a planilha com as ordens de produção
arquivo_excel = 'etiquetas_geradas.xlsx'
ordens_df = pd.read_excel(arquivo_excel)

# Selecionar a ordem de produção
numero_ordem = input("Digite o número da ordem: ")
ordem_selecionada = ordens_df[ordens_df['Ordem'] == int(numero_ordem)]

# Iterar pelas linhas da ordem selecionada
for _, linha in ordem_selecionada.iterrows():
    descricao = linha['Descrição']
    material = linha['Material']
    numero_serie = linha['Número de Série']
    
    # Gera etiqueta visual e ZPL para cada linha
    gerar_etiqueta(descricao, material, numero_serie)
    zpl_code = gerar_zpl(descricao, material, numero_serie)
    
    # Enviar para a impressora
    printer_ip = "192.168.221.99"  # IP da impressora
    printer_port = 9100
    imprimir_zebra(printer_ip, printer_port, zpl_code)

print("Etiquetas geradas e enviadas para a impressora!")
