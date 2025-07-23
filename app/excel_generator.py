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
    'Legal Hold/Product Name': 'N/A', 'Data Owner': 'N/A', 'Notes': ''
}

def create_certificate_workbook(lista_certificados):
    """
    Cria um workbook do Excel em memória a partir de uma lista de dados de certificados. (VERSÃO 3.0)
    """
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(CABECALHO)

    for dados in lista_certificados:
        def get_valor(campo):
            return dados.get(campo, '').upper()

        equipamento = dados.get('equipamento', 'N/A')
        num_certificado = dados.get('num_certificado', 'N/A')

        partes_descricao = [
            f"Certificado de Calibração - Nº: {num_certificado}",
            f"Equipamento: {equipamento}"
        ]

        for campo in ['id', 'tag', 'modelo', 'fabricante', 'sala', 'bloco']:
            valor = get_valor(campo)
            if valor:
                partes_descricao.append(f"{campo.capitalize()}: {valor}")

        data_calibracao_obj = datetime.strptime(dados['data'], '%d/%m/%Y')
        data_formatada_excel = data_calibracao_obj.strftime('%d/%m/%Y')
        data_formatada_descricao = data_calibracao_obj.strftime('%d/%b/%Y').upper()
        partes_descricao.append(data_formatada_descricao)

        descricao = " - ".join(partes_descricao)

        data_destruicao = "31/12/2036"

        tag = get_valor('tag')

        # Montar a linha completa
        linha_completa = [
            dados['barcode'],
            VALORES_FIXOS['Box #/File No/Unique ID'], VALORES_FIXOS['DEPT.'],
            data_formatada_excel,
            data_formatada_excel,
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