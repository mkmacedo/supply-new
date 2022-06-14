from cmath import nan
from datetime import date, datetime, timedelta
#from icecream import ic
import math
from operator import index
#from tracemalloc import start
import pandas as pd
import Regexes
import sys
import numpy as np
import re
from ReadSheets_dev import excelIdentifier
import traceback

month = sys.argv[1]
sheets = sys.argv[2:]

#filenames = sys.stdin.read()
#print(filenames)

#sheets = filenames.split('\n')
#sheets.pop()
print(sheets)


class Medicamentos:
    
    def __init__(self, sheets):
        xlIdentifier = excelIdentifier()
        self.df_vendas = None
        self.df_colocado = None
        self.df_forecast = None
        self.df_produtos = None
        self.df_estoque_blocked = None
        self.df_estoque_all = None
        self.df_parametros = None
        self.df_jda = None
        self.df_drp = None


        for sheet in sheets:

            xlSheet = xlIdentifier.identifySpreadSheet(sheet.strip())

            if xlSheet == 'vendas':
                self.df_vendas = pd.read_excel(sheet.strip())

            elif xlSheet == 'colocado':
                self.df_colocado = pd.read_excel(sheet.strip())
            
            elif xlSheet == 'forecast':
                self.df_forecast = pd.read_excel(sheet.strip())
            
            elif xlSheet == 'produtos':
                self.df_produtos = pd.read_excel(sheet.strip())
            
            elif xlSheet == 'bloqueado':
                temp_df = pd.read_excel(sheet.strip())
                if type(self.df_estoque_blocked) == type(None):
                    self.df_estoque_blocked = temp_df
                elif len(temp_df) <= len(self.df_estoque_blocked):
                    self.df_estoque_all = self.df_estoque_blocked
                    self.df_estoque_blocked = temp_df
                else:
                    self.df_estoque_all = temp_df

            elif xlSheet == 'all':
                temp_df = pd.read_excel(sheet.strip())
                if type(self.df_estoque_all) == type(None):
                    self.df_estoque_all = temp_df
                elif len(temp_df) >= len(self.df_estoque_all):
                    self.df_estoque_blocked = self.df_estoque_all
                    self.df_estoque_all = temp_df
                else:
                    self.df_estoque_blocked = temp_df

            elif xlSheet == 'parametros':
                self.df_parametros = pd.read_excel(sheet.strip())

            elif xlSheet == 'entrada':
                self.df_jda = pd.read_excel(sheet.strip())

            elif xlSheet == 'drp':
                self.df_drp = pd.read_excel(sheet.strip())

            

        #self.df_vendas = pd.read_excel(sheets[0].strip())
        #self.df_produtos = pd.read_excel(sheets[1].strip())
        #self.df_estoque_blocked = pd.read_excel(sheets[2].strip())
        #self.df_estoque_all = pd.read_excel(sheets[3].strip())
        #self.df_forecast = pd.read_excel(sheets[4].strip())
        #self.df_colocado = pd.read_excel(sheets[5].strip())
        #self.df_biotech = pd.read_excel(sheets[6].strip())
        #self.df_jda = pd.read_excel(sheets[7].strip())
        #drp = pd.ExcelFile(sheets[6].strip())
        #self.df_drp = pd.read_excel(drp, 'DRP+SS')
        #self.df_parametros = pd.read_excel('Parâmetros - Supply Planning Biotech.xlsx')
        self.params_dict = {}
        for idx in range(len(self.df_parametros)):
            self.params_dict[self.df_parametros.at[idx, 'Product Code']] = int(self.df_parametros.at[idx, 'Validade mínima para venda (meses)'])
            #print(type(self.df_parametros.at[idx, 'Validade mínima para venda (meses)']))
        #print(self.params_dict)


        self.d = {}

    def calcular(self, month):

        material = self.df_estoque_all['Material No']
        material = material.unique()

        df_provisioning = None
        
        for f in material:

            self.d[f] = {}
            self.d[f]['Sales'] = 0
            self.d[f]['Delivery'] = ''
            self.d[f]['Colocado'] = 0
            self.d[f]['Batch'] = {}

            for i in range(len(self.df_vendas)):
                if str(self.df_vendas.loc[i, 'Material']) == f:
                    self.d[f]['Sales'] += self.df_vendas.loc[i, 'Quantity'] * (-1)

            for i in range(len(self.df_colocado)):
                if str(self.df_colocado.loc[i, 'Código']) == f:# and str(self.df_colocado.loc[i, 'Código']).replace('.','').replace(',','').isdigit():
                    self.d[f]['Colocado'] = self.df_colocado.loc[i, 'Colocado']
            
            #for key in list(self.d.keys()):
            try:
                self.d[f]['Delivery'] = int(self.d[f]['Colocado']) - int(self.d[f]['Sales'])

            except:
                pass
            
            
            for i in range(len(self.df_estoque_all)):

                if str(self.df_estoque_all.loc[i, 'Material No']) == f:

                    if self.d[f].get("Description") == None:
                        self.d[f]['Description'] = self.df_estoque_all.loc[i, 'Material Description']

                    if self.d[f]['Batch'].get(str(self.df_estoque_all.loc[i, 'Batch'])) == None:
                        self.d[f]['Batch'][str(self.df_estoque_all.loc[i, 'Batch'])] = {}
                    
                    if self.d[f]['Batch'][str(self.df_estoque_all.loc[i, 'Batch'])].get('Stock Amount') == None:
                        self.d[f]['Batch'][str(self.df_estoque_all.loc[i, 'Batch'])]['Stock Amount'] = self.df_estoque_all.loc[i, 'Stock']
                    else:
                        self.d[f]['Batch'][str(self.df_estoque_all.loc[i, 'Batch'])]['Stock Amount'] += self.df_estoque_all.loc[i, 'Stock']
                    self.d[f]['Batch'][str(self.df_estoque_all.loc[i, 'Batch'])]['Plant'] = str(self.df_estoque_all.loc[i, 'Plant'])
                    self.d[f]['Batch'][str(self.df_estoque_all.loc[i, 'Batch'])]['Batch status key'] = str(self.df_estoque_all.loc[i, 'Batch status key'])

                    self.d[f]['Batch'][str(self.df_estoque_all.loc[i, 'Batch'])]['Blocked'] = ''
                    for idx in range(len(self.df_estoque_blocked)):
                        if self.df_estoque_blocked.loc[idx, 'Material No'] == f and self.df_estoque_blocked.loc[idx, 'Batch'] == self.df_estoque_all.loc[i, 'Batch']:
                            try:
                                self.d[f]['Batch'][str(self.df_estoque_all.loc[i, 'Batch'])]['Blocked'] = self.df_estoque_blocked.loc[idx, 'Stock']
                            except:
                                ...

                    #Expiration date
                    batchExpDate = (self.df_estoque_all.loc[i, 'Expiration date'], self.df_estoque_all.loc[i, 'Expiration date'].strftime('%Y-%m-%d'))
                    self.d[f]['Batch'][str(self.df_estoque_all.loc[i, 'Batch'])]['Shelf life'] = batchExpDate[1] # Tuple Datetime

                    #days (timedelta)
                    delta = str(date.today() - batchExpDate[0].date())
                    self.d[f]['Batch'][str(self.df_estoque_all.loc[i, 'Batch'])]['Days'] = eval(delta[:delta.find(' days')]) if delta.find(' days') != -1 else eval(delta[:delta.find(' day')]) if delta.find(' day') != -1 else 0
                    self.d[f]['Batch'][str(self.df_estoque_all.loc[i, 'Batch'])]['Month'] = float('{:.1f}'.format(eval(delta[:delta.find(' days')])/30)) if delta.find(' days') != -1 else float('{:.1f}'.format(eval(delta[:delta.find(' day')])/30)) if delta.find(' day') != -1 else 0

                    limit = batchExpDate[0].date() - timedelta(days=30*self.params_dict.get(f, 12))
                    self.d[f]['Batch'][str(self.df_estoque_all.loc[i, 'Batch'])]['Limit sales date'] = (limit, limit.strftime('%Y-%m-%d'))[1] # Tuple Datetime



            #Produtos (Linhas roxas)
            for i in range(len(self.df_produtos)):

                if str(self.df_produtos.loc[i, 'Código']) == f:

                    if self.d[f]['Batch'].get(str(self.df_produtos.loc[i, 'Batch'])) == None:
                        if self.d[f].get('batchAbaProdutos') == None:
                            self.d[f]['batchAbaProdutos'] = {}

                        if self.d[f]['batchAbaProdutos'].get(str(self.df_produtos.loc[i, 'Batch'])) == None:
                            self.d[f]['batchAbaProdutos'][str(self.df_produtos.loc[i, 'Batch'])] = {}
                    
                        if self.d[f]['batchAbaProdutos'][str(self.df_produtos.loc[i, 'Batch'])].get('Stock Amount') == None:
                            self.d[f]['batchAbaProdutos'][str(self.df_produtos.loc[i, 'Batch'])]['Stock Amount'] = self.df_produtos.loc[i, 'Amount']
                        else:
                            self.d[f]['batchAbaProdutos'][str(self.df_produtos.loc[i, 'Batch'])]['Stock Amount'] += self.df_produtos.loc[i, 'Amount']

                        self.d[f]['batchAbaProdutos'][str(self.df_produtos.loc[i, 'Batch'])]['Plant'] = ''
                        self.d[f]['batchAbaProdutos'][str(self.df_produtos.loc[i, 'Batch'])]['Batch status key'] = ''

                        self.d[f]['batchAbaProdutos'][str(self.df_produtos.loc[i, 'Batch'])]['Blocked'] = ''
                        for idx in range(len(self.df_estoque_blocked)):
                            if self.df_estoque_blocked.loc[idx, 'Material No'] == f and self.df_estoque_blocked.loc[idx, 'Batch'] == self.df_produtos.loc[i, 'Batch']:
                                try:
                                    self.d[f]['batchAbaProdutos'][str(self.df_produtos.loc[i, 'Batch'])]['Blocked'] = self.df_estoque_blocked.loc[idx, 'Stock']
                                except:
                                    ...

                        #Expiration date
                        batchAbaProdutosExpDate = (self.df_produtos.loc[i, 'Validade'], self.df_produtos.loc[i, 'Validade'].strftime('%Y-%m-%d')) #Tuple Datetime
                        self.d[f]['batchAbaProdutos'][str(self.df_produtos.loc[i, 'Batch'])]['Shelf life'] = batchAbaProdutosExpDate[1]

                        #days (timedelta)
                        delta = str(date.today() - batchAbaProdutosExpDate[0].date())
                        self.d[f]['batchAbaProdutos'][str(self.df_produtos.loc[i, 'Batch'])]['Days'] = eval(delta[:delta.find(' days')]) if delta.find(' days') != -1 else eval(delta[:delta.find(' day')]) if delta.find(' day') != -1 else 0
                        self.d[f]['batchAbaProdutos'][str(self.df_produtos.loc[i, 'Batch'])]['Month'] = float('{:.1f}'.format(eval(delta[:delta.find(' days')])/30)) if delta.find(' days') != -1 else float('{:.1f}'.format(eval(delta[:delta.find(' day')])/30)) if delta.find(' day') != -1 else 0

                        limit = batchAbaProdutosExpDate[0].date() - timedelta(days=30*self.params_dict.get(f, 12))
                        self.d[f]['batchAbaProdutos'][str(self.df_produtos.loc[i, 'Batch'])]['Limit sales date'] = (limit, limit.strftime('%Y-%m-%d'))[1] #Tuple Datetime


        df = {} # Chave ==>> Código do Material; Valor ==>> DataFrame
        codes = list(self.d.keys())

        for key in codes:
            df[key] = pd.DataFrame()
        
        indexes = list(self.df_forecast.columns) # Lista de colunas da planilha Forecast
        JDA_Cols = list(self.df_jda.columns)

        beginning = 0
        limit = 0
        for i in range(len(indexes)):        
            if indexes[i] == month:
                beginning = i
                limit = min(len(indexes), len(JDA_Cols))


        for i in range(len(self.df_forecast)):

            code = self.df_forecast.loc[i, 'Product Code'] # Código do Material obtido da planilha Forecast

            if code in codes:

                forecast = self.df_forecast.loc[i]

                df[code]['Material No'] = [code] * (limit - beginning)
   
                df[code]['Meses'] = indexes[beginning:limit]

                df[code]['Forecast'] = forecast[list(df[code]['Meses'])].values

                df[code]['Entrada'] = df[code].apply(lambda x: 0., axis = 1)

                df[code]['EstoqueInicial'] = df[code].apply(lambda x: 0., axis = 1)

                df[code]['EstoqueFinal'] = df[code].apply(lambda x: 0., axis = 1)

                df[code]['CoberturaInicial'] = df[code].apply(lambda x: '0.0%', axis = 1)

                df[code]['CoberturaFinal'] = df[code].apply(lambda x: '0.0%', axis = 1)

                df[code]['EstoquePolitica'] = df[code].apply(lambda x: '0.0', axis = 1)
                
                #print(df[code])

        for i in range(len(self.df_jda)):

            if self.df_jda.loc[i,'Projection Columns'] in ['CommitIntransIn', 'ActualIntransIn', 'RecArriv']:

                code = self.df_jda.loc[i, 'Item']
                
                if code in codes:
                    
                    entrada = pd.Series(data=np.zeros((1,len(indexes[beginning:limit])))[0],index=indexes[beginning:limit])

                    startColumn = None
                    for d in JDA_Cols[15:]:
                        r = re.search(r'[0-9][0-9]\.[0-9][0-9]\.[0-9][0-9]', d)
                        rStr = ''
                        if r != None:
                            rStr = r.group().replace(".", "/")
                            dateObj = datetime.strptime(rStr, "%d/%m/%y")
                            rStr = dateObj.strftime("%b %Y").upper()

                        if rStr == month:
                            startColumn = d 
                            break

                    tempSeries = self.df_jda.loc[i]
                    jdaBeginning = JDA_Cols.index(startColumn)
                    jdaLimit = min(len(indexes), len(JDA_Cols))
                    tempSeries = tempSeries[JDA_Cols[jdaBeginning:jdaLimit]]

                    for index, _ in enumerate(entrada):
                        m = re.search(r'[0-9][0-9]\.[0-9][0-9]\.[0-9][0-9]', str(list(tempSeries.index)[index]))#.group()
                        if m != None:
                            m = m.group().replace(".", "/")
                            dateObj_ = datetime.strptime(m, "%d/%m/%y")
                            m = dateObj_.strftime("%b %Y").upper()
                            entrada[m] = tempSeries[list(tempSeries.index)[index]]


                    for index, v in enumerate(df[code]['Entrada']):

                        val = str(entrada[list(entrada.index)[index]])
                        if val.replace(',', '').replace('.', '').replace('-', '').isdigit():
                            val = eval(val.replace(',', ''))

                        if type(val) == type(1.0) or type(val) == type(1):
                            
                            df[code].at[index,'Entrada'] += val
                        #print(df[code].at[index,'Entrada'])
                        #print(index,v)
                    #print(df[code])
        #print(df)

        for key in list(df.keys()): #Percorrendo DataFrames em df
            epFactor = 0 

            for j in range(len(self.df_drp)):
                if self.df_drp.at[j, '*Item'] == key:
                    try:
                        ss_value = self.df_drp.at[j, 'SS (min)']/30 - 1
                        epFactor = ss_value if ss_value > 0 else 0
                    except:
                        #traceback.print_exc()
                        pass


            for k in range(len(self.df_estoque_blocked)):
                if self.df_estoque_blocked.loc[k, 'Material No'] == key and self.df_estoque_blocked.loc[k, 'Batch status key'] == 0:
                    try:
                        df[key].at[0, 'EstoqueInicial'] = self.df_estoque_blocked.loc[k, 'Stock']*(-1)
                    except:
                        ...
            
            for k in range(len(self.df_estoque_all)):
                if key == self.df_estoque_all.loc[k, 'Material No']:
                    #print(self.df_estoque_all.loc[k, 'Stock'])
                    #print(key)

                    for i in range(len(df[key])):
                        
                        if i == 0:# and eInicial[0] == key:
                            if self.d.get(key) != None:# and eInicial[0] == key:

                                if self.df_estoque_all.loc[k, 'Batch status key'] == 0:
                                    df[key].at[i, 'EstoqueInicial'] += self.df_estoque_all.loc[k, 'Stock']#eInicial[1]

                                if df[key].at[i, 'Forecast'] > self.d[key].get('Colocado'):
                                    df[key].at[i, 'EstoqueFinal'] = df[key].at[i, 'EstoqueInicial'] + df[key].at[i, 'Entrada'] - df[key].at[i, 'Forecast']
                                else:
                                    df[key].at[i, 'EstoqueFinal'] = df[key].at[i, 'EstoqueInicial'] + df[key].at[i, 'Entrada'] - self.d[key].get('Colocado')

                                try:
                                    df[key].at[i, 'CoberturaInicial'] =  '{:.2%}'.format(df[key].at[i, 'EstoqueInicial']/df[key].at[i, 'Forecast'])       
                                except:
                                    pass
                                
                                try:
                                    df[key].at[i, 'CoberturaFinal'] = '{:.2%}'.format(df[key].at[i, 'EstoqueFinal'] / df[key].at[i+1, 'Forecast'])

                                except:
                                    pass

                                try:
                                    df[key].at[i, 'EstoquePolitica'] = df[key].at[i, 'Forecast'] + df[key].at[i + 1, 'Forecast']*epFactor
                                except:
                                    pass
                                
                        else:

                            try:
                                df[key].at[i, 'EstoqueInicial'] = df[key].at[i - 1, 'EstoqueFinal'] #mudar
                                df[key].at[i, 'EstoqueFinal'] = df[key].at[i, 'EstoqueInicial'] + df[key].at[i, 'Entrada'] - df[key].at[i, 'Forecast']
                            except:
                                pass

                            try:
                                df[key].at[i, 'CoberturaInicial'] =  '{:.2%}'.format(df[key].at[i, 'EstoqueInicial']/df[key].at[i, 'Forecast'])
                            except:
                                pass

                            try:
                                df[key].at[i, 'CoberturaFinal'] = '{:.2%}'.format(df[key].at[i, 'EstoqueFinal'] / df[key].at[i+1, 'Forecast'])
                            except:
                                pass
                            
                            try:
                                df[key].at[i, 'EstoquePolitica'] = df[key].at[i, 'Forecast'] + df[key].at[i + 1, 'Forecast']*epFactor
                            except:
                                pass
            


            if type(df_provisioning) == type(None):
                df_provisioning = df[key]
            else:
                df_provisioning = df_provisioning.append(df[key], ignore_index=True)

            df_provisioning.to_excel('planlha_supply2.xlsx')


        df_table = None
        for key in list(self.d.keys()):

            #print(key, '--', self.d[key])
            batchList = []
            stockAmountList = []
            plantList = []
            batchStatusKeyList = []
            shelfLifeList = []
            daysList = []
            monthsList = []
            limitSalesDateList = []
            blockedList = []


            for batchKey in list(self.d[key]['Batch']):

                batchList.append(batchKey)

                stockAmount = self.d[key]['Batch'][batchKey].get('Stock Amount')
                stockAmountList.append(stockAmount)

                plant = self.d[key]['Batch'][batchKey].get('Plant')
                plantList.append(plant)

                batchStatusKey = self.d[key]['Batch'][batchKey].get('Batch status key')
                batchStatusKeyList.append(batchStatusKey)

                shelfLife = self.d[key]['Batch'][batchKey].get('Shelf life')
                shelfLifeList.append(shelfLife)

                days = self.d[key]['Batch'][batchKey].get('Days')
                daysList.append(days)

                months = self.d[key]['Batch'][batchKey].get('Month')
                monthsList.append(months)

                limitSalesDate = self.d[key]['Batch'][batchKey].get('Limit sales date')
                limitSalesDateList.append(limitSalesDate)

                blocked = self.d[key]['Batch'][batchKey].get('Blocked')
                blockedList.append(blocked)

            if self.d[key].get('batchAbaProdutos') != None:
                for batchKey in list(self.d[key]['batchAbaProdutos']):

                    batchList.append(batchKey + ' (Trânsito)')

                    stockAmount = self.d[key]['batchAbaProdutos'][batchKey].get('Stock Amount')
                    stockAmountList.append(stockAmount)

                    plant = self.d[key]['batchAbaProdutos'][batchKey].get('Plant')
                    plantList.append(plant)

                    batchStatusKey = self.d[key]['batchAbaProdutos'][batchKey].get('Batch status key')
                    batchStatusKeyList.append(batchStatusKey)

                    shelfLife = self.d[key]['batchAbaProdutos'][batchKey].get('Shelf life')
                    shelfLifeList.append(shelfLife)

                    days = self.d[key]['batchAbaProdutos'][batchKey].get('Days')
                    daysList.append(days)

                    months = self.d[key]['batchAbaProdutos'][batchKey].get('Month')
                    monthsList.append(months)

                    limitSalesDate = self.d[key]['batchAbaProdutos'][batchKey].get('Limit sales date')
                    limitSalesDateList.append(limitSalesDate)

                    blocked = self.d[key]['batchAbaProdutos'][batchKey].get('Blocked')
                    blockedList.append(blocked)

            lgth = len(batchList)
            productList = [key]*lgth
            descriptionList = [self.d[key].get('Description')]*lgth
            salesList = [self.d[key].get('Sales')]*lgth
            deliveryList = [self.d[key].get('Delivery')]*lgth
            colocadoList = [self.d[key].get('Colocado')]*lgth
            destruicaoList = [''] *lgth 
            #blockedList = [0]*lgth

            if len(batchList) != len(stockAmountList) and len(batchList) != len(limitSalesDateList):
                print('FALSE')
            #print()
            #print()

            d = {
                    'Material No': productList,
                    'Description': descriptionList,
                    'Sales': salesList,
                    'Delivery': deliveryList,
                    'Colocado': colocadoList,
                    'Batch': batchList,
                    'Stock Amount': stockAmountList, 
                    'Shelf Life': shelfLifeList,
                    'Days': daysList,
                    'Month': monthsList,
                    'Limit Sales Date': limitSalesDateList,
                    'Plant': plantList,
                    'BSK': batchStatusKeyList,
                    'Blocked': blockedList,
                    'Destruição': destruicaoList 
                    }

            if type(df_table) == type(None):
                df_table = pd.DataFrame(data=d)
            else:
                df_table = df_table.append(pd.DataFrame(data=d), ignore_index=True)


        destruction = {
            'Material No': [''], 
            'Description': [''], 
            'Batch': [''], 
            'Shelf Life': [''], 
            'Month': [''], 
            'Plant': [''], 
            'BSK': ['']
            }
        df_destruction = pd.DataFrame(data=destruction)
        destructionTableCols = ['Material No', 'Description', 'Batch', 'Stock Amount', 'Shelf Life', 'Month', 'Plant', 'BSK']
        for i in range(len(df_table)):
            try:
                if df_table.at[i, 'Month'] >= -6:
                    #print(df_table.at[i, 'Month'])
                    df_destruction = df_destruction.append(df_table.loc[i], ignore_index=True)
                    df_table.at[i, 'Destruição'] = 'DESTRUIR'
            except:
                pass
        df_destruction = df_destruction[destructionTableCols]

        df_table.to_excel('planilha_supply1.xlsx')
        df_destruction.to_excel('planilha_destruicao.xlsx')
        #print(tables)
        #df_final = tables[0]
        #print(df_final.columns)
        #print(tables[1].columns)
        #df_final = df_final.append(tables[1])
        #print(df_final)

        #for t in tables[1:]:
        #    df_final.append(t)
        #print(df_final)



x = Medicamentos(sheets)
x.calcular(month)