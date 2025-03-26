import pandas as pd
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageDraw, ImageFont
import qrcode
import socket
from zpl import ZPLDocument, ZPLField

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

# Função para verificar se a etiqueta já foi impressa
def etiqueta_ja_impressa(ordem, material, numero_serie):
    log_file = 'log_impressao.csv'
    
    try:
        log_df = pd.read_csv(log_file)
    except FileNotFoundError:
        log_df = pd.DataFrame(columns=['Ordem', 'Material', 'Número de Série'])
    
    # Verificar se já existe a entrada no log
    if ((log_df['Ordem'] == ordem) & (log_df['Material'] == material) & (log_df['Número de Série'] == numero_serie)).any():
        return True
    else:
        return False

# Função para registrar a impressão no log
def registrar_impressao(ordem, material, numero_serie):
    log_file = 'log_impressao.csv'
    
    try:
        log_df = pd.read_csv(log_file)
    except FileNotFoundError:
        log_df = pd.DataFrame(columns=['Ordem', 'Material', 'Número de Série'])
    
    new_entry = pd.DataFrame({'Ordem': [ordem], 'Material': [material], 'Número de Série': [numero_serie]})
    
    # Usando pd.concat para adicionar a nova entrada
    log_df = pd.concat([log_df, new_entry], ignore_index=True)
    log_df.to_csv(log_file, index=False)


# Função chamada quando o botão de imprimir for clicado
def imprimir_etiquetas():
    try:
        numero_ordem = int(entry_ordem.get())  # Pega o número da ordem do campo de entrada
    except ValueError:
        messagebox.showerror("Erro", "Por favor, insira um número de ordem válido.")
        return
    
    # Carregar a planilha com as ordens de produção
    arquivo_excel = 'etiquetas_geradas.xlsx'
    ordens_df = pd.read_excel(arquivo_excel)
    
    ordem_selecionada = ordens_df[ordens_df['Ordem'] == numero_ordem]
    
    if ordem_selecionada.empty:
        messagebox.showerror("Erro", "Ordem não encontrada.")
        return
    
    # Iterar pelas linhas da ordem selecionada
    for _, linha in ordem_selecionada.iterrows():
        descricao = linha['Descrição']
        material = linha['Material']
        numero_serie = linha['Número de Série']
        
        if etiqueta_ja_impressa(numero_ordem, material, numero_serie):
            messagebox.showinfo("Aviso", f"Etiqueta para o material {material} já foi impressa.")
            continue
        
        # Gera etiqueta visual e ZPL para cada linha
        gerar_etiqueta(descricao, material, numero_serie)
        zpl_code = gerar_zpl(descricao, material, numero_serie)
        
        # Enviar para a impressora
        printer_ip = "192.168.221.99"  # IP da impressora
        printer_port = 9100
        imprimir_zebra(printer_ip, printer_port, zpl_code)
        
        # Registrar a impressão no log
        registrar_impressao(numero_ordem, material, numero_serie)
    
    messagebox.showinfo("Sucesso", "Etiquetas geradas e enviadas para a impressora!")

# Criar a interface com Tkinter
root = tk.Tk()
root.title("Sistema de Impressão de Etiquetas")

# Tamanho da janela
root.geometry("400x200")

# Campo de entrada para número da ordem
label_ordem = tk.Label(root, text="Número da Ordem:")
label_ordem.pack(pady=10)

entry_ordem = tk.Entry(root, width=30)
entry_ordem.pack()

# Botão para imprimir etiquetas
button_imprimir = tk.Button(root, text="Imprimir Etiquetas", command=imprimir_etiquetas)
button_imprimir.pack(pady=20)

# Rodando a interface
root.mainloop()
