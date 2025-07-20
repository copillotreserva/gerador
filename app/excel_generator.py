import openpyxl
from datetime import datetime
from io import BytesIO

# --- CABECALHO E VALORES FIXOS (SEM MUDANÇAS) ---
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
    Cria um workbook do Excel em memória a partir de uma lista de dados de certificados.
    """
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(CABECALHO)

    for dados in lista_certificados:
        # Pega os valores dos campos. Se o campo estiver vazio, usa 'N/A'.
        tag = dados.get('tag', '') or 'N/A'
        sala = dados.get('sala', '') or 'N/A'
        modelo = dados.get('modelo', '') or 'N/A'
        fabricante = dados.get('fabricante', '') or 'N/A'

        # 1. Processar a data
        data_calibracao_obj = datetime.strptime(dados['data'], '%d/%m/%Y')
        data_formatada_excel = data_calibracao_obj.strftime('%d-%b-%Y')
        data_formatada_descricao = data_calibracao_obj.strftime('%d/%b/%Y')

        # 2. Calcular a Data de Destruição
        ano_destruicao = data_calibracao_obj.year + 22
        data_destruicao = f"31/Dec/{ano_destruicao}"

        # 3. Montar a Descrição
        descricao = (
            f"Certificado de Calibração N° {dados['num_certificado']}, "
            f"{dados['tipo_instrumento']}, TAG: {tag} - "
            f"Sala: {sala} - Modelo: {modelo}, "
            f"Fabricante: {fabricante} - {data_formatada_descricao}"
        )

        # 4. Montar a linha completa
        linha_completa = [
            dados['barcode'],
            VALORES_FIXOS['Box #/File No/Unique ID'], VALORES_FIXOS['DEPT.'],
            data_formatada_excel, data_formatada_excel, VALORES_FIXOS['Record Series Code'],
            VALORES_FIXOS['CATEGORY'], VALORES_FIXOS['SUBCATEGORY'], VALORES_FIXOS['Record Series Title/Type'],
            descricao, tag, VALORES_FIXOS['Retention Period'], data_destruicao,
            VALORES_FIXOS['Legal Hold/Product Name'], VALORES_FIXOS['Data Owner'], VALORES_FIXOS['Notes']
        ]
        sheet.append(linha_completa)

    memoria_excel = BytesIO()
    workbook.save(memoria_excel)
    memoria_excel.seek(0)

    return memoria_excel
