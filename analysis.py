import pandas as pd
import numpy as np
import os, glob

data = input('Nombre archivo:')
monto = float(input('Monta base:'))
variacion = float(input('Variación:'))

def contracts(data, monto, variacion):     
    df = pd.read_excel(( os.getcwd() + '/data/' + data ).replace('\\','/'), header=7)
    df.drop(df.columns[[1,2,4,5,6,7,8]], axis=1, inplace=True)
    df.columns = ['Contrato', 'Descripción', 'Total']
    df.dropna(subset=['Total'], how='any', inplace=True)

    contrato = (df[df['Contrato'].str.contains('nombre', case=False)].iloc[:,0]).reset_index(drop=True)
    base = (df[df['Contrato'].str.contains('revisión 0.0', case=False)].iloc[:,2]).reset_index(drop=True)
    cierre = (df[df['Contrato'].str.contains('total compro', case=False)].iloc[:,2]).reset_index(drop=True)

    df1 = pd.concat([contrato, base, cierre], axis=1)
    df1.columns = ['Contrato', 'Base', 'Cierre']
    df1['Delta'] = df1['Cierre'] - df1['Base']
    df1['Contrato'] = df1.Contrato.str.slice(start=8) #Remueve los primeros 8 caracteres del string en la columna "Contratos".
    df1['Variación'] = df1['Delta']/(df1['Base'] + 0.00000001)

    df2 = df1[df1.Delta > 0].reset_index(drop=True)
    
    df3 = (df2[(df2['Base'] > monto) & (df2['Variación'] > variacion)]).sort_values(by=['Base'], ascending=False).reset_index(drop=True)

    return df, df1, df2, df3

df, df1, df2, df3 = contracts(data, monto, variacion)

def contract_analysis_base(df, monto, variacion):

    dfs = df[df.Contrato.str.match('0.0|nombre|revisión 0.0|total compro', case=False)].reset_index(drop=True)
    dfs[dfs['Descripción'].str.contains('suministr', case=False)==True]

    ind_nombre = dfs[dfs['Contrato'].str.contains('nombre', case=False)].index.tolist()
    ind_base = dfs[dfs['Contrato'].str.contains('revisión 0.0', case=False)].index.tolist()
    ind_cierre = dfs[dfs['Contrato'].str.contains('total compro', case=False)].index.tolist()
    ind_df = pd.DataFrame(np.column_stack([ind_nombre, ind_base, ind_cierre]), columns=['Índice contrato', 'Índice base', 'Índice cierre'])

    l = []
    for i in range(len(ind_df)):
        if (dfs.Total[ind_df['Índice base'][i]] > monto) and (dfs.Total[ind_df['Índice cierre'][i]]/dfs.Total[ind_df['Índice base'][i]] - 1 > variacion):
            s = dfs.iloc[ind_df['Índice contrato'][i]:ind_df['Índice base'][i], :]
            l.append(s)
            dfx = pd.concat(l, ignore_index=True) 
        
    return dfx

df_base = contract_analysis_base(df, monto, variacion)

def contract_selection(df, monto, variacion):

    df4 = (df[~df['Contrato'].str.match('0.0|totales fina', case=False)]).reset_index(drop=True) 

    ind_nombre = df4[df4['Contrato'].str.contains('nombre', case=False)].index.tolist()
    ind_base = df4[df4['Contrato'].str.contains('revisión 0.0', case=False)].index.tolist()
    ind_cierre = df4[df4['Contrato'].str.contains('total comprom', case=False)].index.tolist()
    ind_df = pd.DataFrame(np.column_stack([ind_nombre, ind_base, ind_cierre]), columns=['Índice contrato', 'Índice base', 'Índice cierre'])

    l = []
    for i in range(len(ind_df)):
        if (df4.Total[ind_df['Índice base'][i]] > monto) and (df4.Total[ind_df['Índice cierre'][i]]/df4.Total[ind_df['Índice base'][i]] - 1 > variacion):
            s = df4.iloc[ind_df['Índice contrato'][i]:ind_df['Índice cierre'][i]+1, :]
            l.append(s)
            df5 = pd.concat(l, ignore_index=True) 

    df5['Contrato'] = df5.Contrato.str.replace('Revisión 0.0 - Totales', 'Costo Base', regex=False)
    df5['Contrato'] = df5.Contrato.str.replace('Total Compromiso', 'Costo Final', regex=False)
    df5 = df5[~df5.Contrato.str.contains('revisión', case=False)]
    return df5

df_s = contract_selection(df, monto, variacion)

resumen = df_s[df_s.Contrato.str.contains('nombre|costo base|costo final', case=False)]
del resumen['Descripción']

with pd.ExcelWriter('selected_contracts_(from code).xlsx') as writer:
    df1.to_excel(writer, sheet_name='Reagrupación', index=False)
    df2.to_excel(writer, sheet_name='Sólo crecimientos', index=False)
    df3.to_excel(writer, sheet_name='Seleccionados', index=False)
    df_s.to_excel(writer, sheet_name='Detalle seleccionados', index=False)
    resumen.to_excel(writer, sheet_name='Resumen seleccionados', index=False)

items = ['extensión plazo', 'obras adicionales', 'materiales', 'ingeniería']

key_words = ['extensión plazo|plazo', 'montaje|obras adicio|adicional|extraordinaria|obraexcav|civil|rellen|reempla|reparaci|instalaci', 'hormig|piping|cañer|tuber|válv|valv|pern|sumin|adqui|estruc|acero', 'ingenieri|cambios alcan|ingenieria de terreno|ingeniería de terreno'] 

def cost_deviations(df, key_words):
    l = []
    for i in range(len(key_words)):
        a = df[df['Descripción'].str.contains(key_words[i], case=False)==True] 
        l.append(a)
        df = df.merge(l[i].drop_duplicates(), on=['Contrato', 'Descripción', 'Total'], how='left', indicator=True)
        df = df[df['_merge'] =='left_only']
        del df['_merge']

    return l

l = cost_deviations(df_s, key_words)

def resumen_items(l, items):
    d = {}
    a = 0 
    for i in range(len(items)):       
            d[items[i]] = l[i].Total.sum()
            a += l[i].Total.sum()
            
    d['Total'] = a
    dff = pd.DataFrame(d, index=['USD', 'Var']).T
    dff.Var = dff.USD/df2.Base.sum()

    return dff

dff = resumen_items(l, items)

#--------------------------------------------------------
dx = resumen_items(cost_deviations(df_base, key_words), items)

with pd.ExcelWriter('cost_deviations_(from code).xlsx') as writer:
    dff.to_excel(writer, sheet_name='Resumen seleccionados', index=True)
    dx.to_excel(writer, sheet_name='Resumen seleccionados base', index=True)
    for i in range(len(items)):
        pd.DataFrame(l[i]).to_excel(writer, sheet_name=items[i], index=False)
