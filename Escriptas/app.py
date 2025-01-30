from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
import os
import xml.etree.ElementTree as ET
import logging

app = Flask(__name__)

# Configurações para upload de arquivos
UPLOAD_FOLDER = 'uploads/'
ALLOWED_EXTENSIONS = {'xml'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Verifica se o diretório de upload existe, se não, cria
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Configuração de logging
logging.basicConfig(level=logging.DEBUG)

# Função para verificar a extensão do arquivo
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Função para processar os dados do XML
def processar_nfe(xml_data):
    try:
        root = ET.fromstring(xml_data)
        
        namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
        
        nNF = root.find('.//nfe:ide/nfe:nNF', namespaces)
        if nNF is not None:
            nNF = nNF.text
        else:
            return {'error': 'Campo nNF não encontrado no XML'}
        
        infCpl = root.find('.//nfe:infAdic/nfe:infCpl', namespaces)
        if infCpl is not None and '@REF ' in infCpl.text:
            call_number = infCpl.text.split('@REF ')[1].split(';')[0]
        else:
            return {'error': 'Call Number não encontrado no XML'}
        
        dados_nfe = {
            'Número da NFe': nNF,
            'Call Number': call_number,
            'Volume Total': 0,
            'Part Numbers': [],
            'Condição das Peças': []
        }

        for det in root.findall('.//nfe:det', namespaces):
            part_number = det.find('.//nfe:prod/nfe:cProd', namespaces)
            volume = det.find('.//nfe:prod/nfe:qCom', namespaces)
            infCpl = root.find('.//nfe:infAdic/nfe:infCpl', namespaces)
            
            if part_number is not None and volume is not