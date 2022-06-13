import pandas as pd

class excelIdentifier:
    def __init__(self):
        self.cols_dictionary = {
            'vendas': ['Material', 'Material description', 'Amt.in loc.cur.', 'Currency', 'Movement Type', 'Plant', 'Storage location', 'Posting Date', 'Quantity', 'Unit of Entry', 'Company Code', 'Purchase order', 'Vendor', 'Batch', 'Entry date', 'Movement Type Text', 'Movement indicator', 'Material Document', 'Order', 'Base Unit of Measure', 'Reference', 'Customer', 'Document Date', 'Material Doc. Year', 'Document Header Text'],
            'colocado': ['Código', 'Descrição', 'Colocado', '% Forecast', 'Forecast', 'Faturado', 'Disp', 'UFD', 'UFD (EURO)', 'Situação', 'Planejador', 'OBS'],
            'forecast': ['Product Code', 'Product Desc', 'Business Line Descr', 'Key Figure'],
            'produtos': ['Status', 'DIVISAO', 'Tipo de \nTransporte', 'Código', 'Product', 'Batch', 'Amount', 'Proforma', 'Purchase Order', 'IA', 'Valor Unitário', 'Valor TOTAL', 'Documentação Completa', 'Validade', 'Dias', 'Meses', 'RMSL', 'Recebimento do \nPré-Alerta', 'Embarque / ETD ATUSA ', 'Envio \nDossie DMS', 'LT envio Dossie', 'Invoice', 'AWB/BL', 'Quantidade de volume', 'Número\nVAQ-Tainer/\nLifeConex', 'Número LI', 'Emissão LI', 'LT Registro LI', 'CHEGADA - ETA', 'Transit Time \nAir - 3 \nSea - 7', 'Mantra Visado', 'Protocolo\nVISCOMEX', 'Recebimento Protocolo', 'Processo ANVISA', 'LT protocolo DMS\n(3 dias)', 'LT Protocolo Anvisa\n(1 dia)', 'Em Exigência?', 'Deferimento\n(7 dias) ', 'LT Inspeçao ANVISA (7 dias)', 'DI', 'Registro', 'Data Desembaraço', 'LT RF\n(1/2 dias)', 'Canal', 'Chegada Merck', 'LT Total Desembaraço', 'LT Chegada Brasil até Entrega na Merck', 'Entrada Protocolo\nBTG', 'LT Entrada Protocolo\n(5 dias)', 'Recebimento \nBTG', 'LT Rec BTG Anvisa\n(7 dias)', 'Previsão Liberação Lote', 'LT Liberação do Lote\n(2 dias)', 'LT Geral BTG', 'LT Chegada-BTG', 'Observação'],
            'bloqueado': ['Material No', 'Material Description', 'Batch', 'Expiration date', 'Batch status key', 'Stock', 'Price', 'Price unit', 'Stock Valuve in LC', 'Plant', 'Storage location', 'Vendor Batch'],
            'all': ['Material No', 'Material Description', 'Batch', 'Expiration date', 'Batch status key', 'Stock', 'Price', 'Price unit', 'Stock Valuve in LC', 'Plant', 'Storage location', 'Vendor Batch'],
            'parametros': ['Product Code', 'Product Desc', 'Corte de validade para destruição (meses)', 'Validade mínima para venda (meses)'],
            'entrada': ['Item', 'Loc', 'Item.1', 'Issue_Manuf', 'Issue_CMG', 'Local code', 'Item descr', 'DRPCovDur', 'MinDRPQty', 'IncDRPQty', 'ProdMinQty', ' ProdIncQty', ' ProdMaxQty', 'SSCov', 'Projection Columns'],
            'drp': ['*Item', 'ItemDescr', '*Loc', 'DRPCovDur (+SS = max)', 'SS (min)']
        }


    def identifySpreadSheet(self, filename):
        df = pd.read_excel(filename)
        for key in list(self.cols_dictionary.keys()):
            if key == 'forecast':
                if list(df.columns)[:4] == self.cols_dictionary[key]:
                    return key
            if key == 'drp':
                if list(df.columns)[:5] == self.cols_dictionary[key]:
                    return key
            if key == 'entrada':
                if list(df.columns)[:15] == self.cols_dictionary[key]:
                    return key
            elif list(df.columns) == self.cols_dictionary[key]:
                return key



#x = excelIdentifier()
#res = x.identifySpreadSheet('DRP.xlsx')
#print(res)

