from PIL import Image, ImageDraw, ImageFont
import qrcode
import socket

# Função para gerar o QR Code
def gerar_qr_code(conteudo):
    qr = qrcode.QRCode(
        version=1,  # Tamanho do QR Code
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=15,
        border=4,
    )
    qr.add_data(conteudo)
    qr.make(fit=True)
    img_qr = qr.make_image(fill='black', back_color='white')
    return img_qr

# Função para gerar a etiqueta
def gerar_etiqueta(material, numero_serie, descricao):
    largura = 1040  # 104mm * 10 (escala para pixels)
    altura = 400    # 40mm * 10 (escala para pixels)
    
    # Criar a imagem de fundo da etiqueta
    etiqueta = Image.new('RGB', (largura, altura), color='white')
    draw = ImageDraw.Draw(etiqueta)
    
    # Fonte para o texto (pode ser substituída por um arquivo .ttf se necessário)
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 50)
    except IOError:
        font = ImageFont.load_default()
    
    # Texto para Material, Número de Série e Descrição
    texto_material = f"Material: {material}"
    texto_serie = f"Nº Série: {numero_serie}"
    texto_descricao = f"Descrição: {descricao}"

    # Gerar o QR Code
    qr_conteudo = f"{material};{numero_serie}"
    qr_img = gerar_qr_code(qr_conteudo)

    # Definir posições do texto e QR Code
    draw.text((20, 20), texto_material, font=font, fill='black')
    draw.text((20, 200), texto_serie, font=font, fill='black')
    draw.text((20, 300), texto_descricao, font=font, fill='black')

    # Colocar o QR Code na etiqueta
    etiqueta.paste(qr_img, (750, 100))

    # Salvar a etiqueta em um arquivo de imagem
    etiqueta.save(f"etiqueta_{numero_serie}.png")
    etiqueta.show()

# Exemplo de uso
material = "717803"
numero_serie = "EVC0333"
descricao = "Caixa Balanc. BCH E115EP3RB03"

gerar_etiqueta(descricao, material, numero_serie,)



def gerar_zpl(descricao, material, numero_serie, qr_code_path):
    # Definindo o layout ZPL
    zpl = "^XA\n"  # Início do formato ZPL
    
    # Tamanho da etiqueta (104mm x 40mm)
    zpl += "^FO20,200^A0N,60,100^FDMaterial: {0}^FS\n".format(material)
    zpl += "^FO20,300^A0N,60,40^FDNº Série: {0}^FS\n".format(numero_serie)
    zpl += "^FO20,20^A0N,60,40^FDDescrição: {0}^FS\n".format(descricao)
    
    # Inserindo o QR Code
    zpl += "^FO750,100^BQN,2,10^FDMA,{0}^FS\n".format(f"{material};{numero_serie}")
    
    # Finalizando o comando ZPL
    zpl += "^XZ\n"  # Fim do formato ZPL

    return zpl

# Exemplo de uso
zpl_code = gerar_zpl(descricao, material, numero_serie, "path_to_qr_code_image.png")

# Salvar o código ZPL em um arquivo (ou enviar para a impressora)
with open("etiqueta.zpl", "w") as file:
    file.write(zpl_code)

# Endereço IP e porta da impressora Zebra
printer_ip = "192.168.221.99"
printer_port = 9100

# Conectar à impressora
sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
sock.connect((printer_ip, printer_port))

# Enviar o código ZPL
sock.sendall(zpl_code.encode())

# Fechar a conexão
sock.close()