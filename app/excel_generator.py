import openpyxl
from datetime import datetime
from io import BytesIO

# --- CABECALHO E VALORES FIXOS ---
CABECALHO = [
    'Barcode #', 'Box #/File No/Unique ID', 'DEPT.', 'Date From', 'Date To',
    'Record Series Code', 'CATEGORY', 'SUBCATEGORY', 'Record Series Title/Type',
    'Record Description or Box Description', 'TAG', 'Retention Period',
    'Destruction Date', 'Legal Hold/Product Name', 'Data Owner', 'Notes'
]

VALORES_FIXOS = {
    'Box #/File No/Unique ID': 'NA', 'DEPT.': 'Engenharia',
    'Record Series Code': 'MAN-007-002', 'CATEGORY': 'Manufacturing', 'SUBCATEGORY': 'Plant Model',
    'Record Series Title/Type': 'Validation lifecycle documentation', 'Retention Period': 'Retired + 11 years',
    'Legal Hold/Product Name': 'N/A', 'Data Owner': 'N/A', 'Notes': 'BLOCO Q'
}

def create_certificate_workbook(lista_certificados):
    """
    Cria um workbook do Excel em memória a partir de uma lista de dados de certificados. (VERSÃO 2.0)
    """
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(CABECALHO)

    for dados in lista_certificados:
        def get_valor(campo):
            return dados.get(campo, '') or 'N/A'

        tag = get_valor('tag')
        sala = get_valor('sala')
        bloco = get_valor('bloco')
        modelo = get_valor('modelo')
        fabricante = get_valor('fabricante')

        # Processar a data
        data_calibracao_obj = datetime.strptime(dados['data'], '%d/%m/%Y')
        # 2. NOVO FORMATO DE DATA PARA AS COLUNAS D e E
        data_formatada_excel = data_calibracao_obj.strftime('%d/%m/%Y')
        data_formatada_descricao = data_calibracao_obj.strftime('%d-%b-%Y') # Mantém formato 'dd-Mês-yyyy' para descrição

        # Calcular a Data de Destruição
        ano_destruicao = data_calibracao_obj.year + 22
        data_destruicao = f"31/Dec/{ano_destruicao}"

        # 3. MUDANÇA NA DESCRIÇÃO (SEPARADORES E CAMPO BLOCO)
        descricao = (
            f"Certificado de Calibração N° {dados['num_certificado']} - "
            f"Tipo: {dados['tipo_instrumento']} - TAG: {tag} - "
            f"Sala: {sala} - Bloco: {bloco} - Modelo: {modelo} - "
            f"Fabricante: {fabricante} - {data_formatada_descricao}"
        )

        # Montar a linha completa
        linha_completa = [
            dados['barcode'],
            VALORES_FIXOS['Box #/File No/Unique ID'], VALORES_FIXOS['DEPT.'],
            data_formatada_excel, # Novo formato aqui
            data_formatada_excel, # Novo formato aqui
            VALORES_FIXOS['Record Series Code'],
            VALORES_FIXOS['CATEGORY'], VALORES_FIXOS['SUBCATEGORY'], VALORES_FIXOS['Record Series Title/Type'],
            descricao, tag, VALORES_FIXOS['Retention Period'], data_destruicao,
            VALORES_FIXOS['Legal Hold/Product Name'], VALORES_FIXOS['Data Owner'], VALORES_FIXOS['Notes']
        ]
        sheet.append(linha_completa)

    memoria_excel = BytesIO()
    workbook.save(memoria_excel)
    memoria_excel.seek(0)

    return memoria_excel