from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
from print import gerar_etiqueta, imprimir_zebra, registrar_impressao

app = Flask(__name__)

@app.route('/')
def index():
    # Carregar as ordens de produção do arquivo Excel
    ordens_df = pd.read_excel('etiquetas_geradas.xlsx')
    ordens = ordens_df['Ordem'].unique()  # Exibir ordens únicas
    return render_template('index.html', ordens=ordens)

@app.route('/imprimir', methods=['POST'])
def imprimir_etiquetas():
    numero_ordem = request.form.get('numero_ordem')
    ordens_df = pd.read_excel('etiquetas_geradas.xlsx')
    
    # Filtrar a ordem selecionada
    ordem_selecionada = ordens_df[ordens_df['Ordem'] == int(numero_ordem)]
    
    # Iterar e imprimir as etiquetas
    for _, linha in ordem_selecionada.iterrows():
        descricao = linha['Descrição']
        material = linha['Material']
        numero_serie = linha['Número de Série']
        
        # Gerar etiqueta e imprimir
        gerar_etiqueta(descricao, material, numero_serie)
        zpl_code = gerar_zpl(descricao, material, numero_serie)
        
        # Enviar para a impressora
        printer_ip = "192.168.221.99"  # IP da impressora
        printer_port = 9100
        imprimir_zebra(printer_ip, printer_port, zpl_code)
        
        # Registrar no log
        registrar_impressao(numero_ordem, material, numero_serie)
    
    return redirect(url_for('index'))  # Redirecionar de volta à página principal

if __name__ == '__main__':
    app.run(debug=True)
