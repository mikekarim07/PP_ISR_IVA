import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import streamlit as st
import plotly.express as px
import base64
import io
from io import StringIO, BytesIO
from streamlit_option_menu import option_menu
#importing the os module
import os
import codecs

#to get the current working directory
# directory = os.getcwd()



st.set_page_config(page_title='TAX - Pago Provisional')



image_url = 'https://github.com/mikekarim07/PP_ISR_IVA/blob/main/kor_logo.png'
st.image("kor_logo_web.png", width=200)
st.title('Cálculo del Pago Provisional 📈')
st.subheader('Cargar los siguientes archivos: Auxiliar, Balanza y Customer del periodo')
# st.write(directory)
st.write("Streamlit version:", st.__version__)
#

Cat_ctas_uploaded_file = st.file_uploader('Selecciona el Archivo que contiene el Catalogo de Cuentas', type='xlsx')
if Cat_ctas_uploaded_file:
    st.markdown('---')
    Catalogo = pd.read_excel(Cat_ctas_uploaded_file, engine='openpyxl', sheet_name='Catalogo',
                            dtype = {'Cuenta': str, 'Descripcion': str, 'Tipo': str,})
    coeficientes = pd.read_excel(Cat_ctas_uploaded_file, engine='openpyxl', sheet_name='CU',
                                 dtype = {'CoCode': str,})

# st.caption('Cargar el Auxiliar del periodo')
Auxiliar_uploaded_file = st.file_uploader('Selecciona el Archivo que contiene el auxiliar del periodo', type='xlsx')
if Auxiliar_uploaded_file:
    st.markdown('---')
    Auxiliar = pd.read_excel(Auxiliar_uploaded_file, engine='openpyxl', skiprows=5, 
                            dtype = {'Period': str, 'Account': str, 'DocumentNo': str, 'Document Header Text': str,
                                'Cost Ctr': str, 'Assignment': str,})

Balanza_uploaded_file = st.file_uploader('Selecciona el Archivo que contiene la Balanza', type='txt')
if Balanza_uploaded_file:
    st.markdown('---')
    Balanza = pd.read_csv(Balanza_uploaded_file, sep='|', header=None,
                          names=['Account', 'Description', 'Saldo Inicial', 'Cargos', 'Abonos', 'Saldo Final', 'CoCode'], encoding='iso-8859-1',
                          dtype = {'Account': str, 'Description': str, 'CoCode': str,})

Customer_uploaded_file = st.file_uploader('Selecciona el Archivo que contiene el customer del periodo', type='xlsx')
if Customer_uploaded_file:
    st.markdown('---')
    Customer = pd.read_excel(Customer_uploaded_file, engine='openpyxl',
                        names=['Varios', '*', 'St', 'Assignment', 'Nombres', 'DocumentNo', 'Typ', 'LCurr', 'Clrng doc.', 'Tx',
                              'Doc. Date', 'Reference', 'Text', 'Amt in loc.cur.', 'Customer'],
                        index_col=None)
    st.subheader('Catalogo de Cuentas')
    st.dataframe(Catalogo)
    
    Auxiliar = Auxiliar.rename(columns={"CoCd": "CoCode", "Amount in local cur.": "Monto"})
    Auxiliar = Auxiliar.dropna(subset=['Account'])
    Auxiliar = Auxiliar[Auxiliar['Monto']<0]
    Auxiliar = Auxiliar.groupby(by=['Account', 'CoCode'], as_index=False)['Monto'].sum()
    Auxiliar['Monto'] = Auxiliar['Monto'].abs()
    Auxiliar['Source'] = 'Auxiliar'
    st.subheader('Auxiliar')
    st.write(Auxiliar.shape)
    st.dataframe(Auxiliar)
    
    st.divider()

    st.subheader('Balanza')
    Balanza['Monto'] = Balanza['Saldo Final'] - Balanza['Saldo Inicial']
    Balanza['Source'] = 'Balanza'
    Balanza[['Account']] = Balanza[['Account']].apply(pd.to_numeric)
    Balanza = Balanza[Balanza['Account']>2000000000]
    Balanza[['Account']] = Balanza[['Account']].astype('string')
    Balanza = Balanza[Balanza['Monto']<0]
    Balanza['Monto'] = Balanza['Monto'].abs()
    st.write(Balanza.shape)
    st.dataframe(Balanza)
    
    st.divider()
    
    def company_code(row):
        if row['Varios'] == ' Company Code':
            return row['Nombres']
        else:
            return None

    def customername(row):
        if row['Varios'] == ' Name':
            return row['Nombres']
        else:
            return None



    Customer['CoCode'] = Customer.apply(company_code, axis=1)    
    Customer['CustomerName'] = Customer.apply(customername, axis=1)
    Customer['CoCode'].fillna(method='ffill', inplace=True)
    Customer['CustomerName'].fillna(method='ffill', inplace=True)
    Customer = Customer.dropna(subset=['Typ'])
    Customer = Customer[Customer['Typ']!='DZ']
    Customer = Customer[Customer['Customer']!='Customer']
    Customer['Tx'].fillna('OK', inplace=True)
    Customer = Customer[(Customer['Tx'] == 'EG') | (Customer['Tx'] == 'OK')]
    Customer = Customer[(Customer['CoCode'] == 'ABMX') | (Customer['CoCode'] == 'AFMX') | (Customer['CoCode'] == 'AHMX')]
    Customer = Customer.groupby(by=['Customer', 'Tx', 'CoCode'], as_index=False)['Amt in loc.cur.'].sum()
    
    def montosiniva(row):
        if row['Tx'] == 'EG':
            return row['Amt in loc.cur.']/1.16
        else:
            return None
    
    Customer['Monto'] = Customer.apply(montosiniva, axis=1)    
    Customer = Customer.groupby(by=['CoCode', 'Customer'], as_index=False)['Monto'].sum()
    Customer = Customer[Customer['Monto']>0]
    Customer = Customer.rename(columns={"Customer": "Account"})
    
    st.subheader('Customer')
    st.write(Customer.shape)
    st.dataframe(Customer)
    
    st.divider()

    
    aux_bal = pd.concat([Auxiliar, Balanza])
    aux_bal = aux_bal.merge(Catalogo, left_on='Account', right_on='Cuenta', how='left')
    aux_bal = aux_bal[aux_bal['Tipo']!='No Aplica']
    # aux_bal = tipo(aux_bal)
    aux_bal['new'] = np.where((aux_bal['Source'] == aux_bal['Tipo']), 'ok', 'no')
    aux_bal = aux_bal[aux_bal['new']=='ok']
    alldata = pd.concat([aux_bal,Customer])
    # alldata = alldata[['Account', 'CoCode', 'Source', 'Descripcion', 'Monto']]
    st.subheader('Datos consolidados')
    st.write(alldata.shape)
    st.dataframe(alldata)
    st.divider()
    summary = alldata.groupby(by=['CoCode'], as_index=False)['Monto'].sum()
    
    st.subheader('Resumen')
    st.write(summary.shape)
    st.dataframe(summary)
    st.divider()
    
    coeficientes = coeficientes[['CoCode', 'Enero']]
    st.subheader('Coeficientes de Utilidad')
    st.write(coeficientes.shape)
    st.dataframe(coeficientes)
    st.divider()
    
    pago_prov = summary.merge(coeficientes, left_on='CoCode', right_on='CoCode', how='left')
    pago_prov['Utilidad Fiscal'] = pago_prov['Monto'] * pago_prov['Enero']
    pago_prov['Tasa ISR'] = .30
    pago_prov['ISR'] = pago_prov['Utilidad Fiscal'] * pago_prov['Tasa ISR']
    
    st.subheader('Calculo de Pago Provisional')
    st.write(pago_prov.shape)
    st.dataframe(pago_prov)
    st.divider()
    
    buffer = io.BytesIO()

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write each dataframe to a different worksheet.
        pago_prov.to_excel(writer, sheet_name='Pago Provisional')
        alldata.to_excel(writer, sheet_name='Consolidado')
        Auxiliar.to_excel(writer, sheet_name='Auxiliar')
        Balanza.to_excel(writer, sheet_name='Balanza')
        Customer.to_excel(writer, sheet_name='Customer')

        writer.save()

    # Set up download link
    b64 = base64.b64encode(buffer.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="Pago_provisional.xlsx">Download Excel File</a>'
    st.markdown(href, unsafe_allow_html=True)

    # Nuevo dataframe unicamente con las columnas de Account en FBL3N y GL_Account en Parametros
    # si va FBL3N_ctas = df_FBL3N[['Account']].astype(str)
    # si va Parametros_ctas = df_parametros[['GL_Account']].astype(str)
    # Eliminar duplicados de FBL3N
    # si va Ctas_unicas = FBL3N_ctas[['Account']].drop_duplicates()
    # Genera un nuevo Dataframe donde se hace el merge de FBL3N y Parametros
    # si va result = pd.merge(Ctas_unicas, Parametros_ctas, left_on = 'Account', right_on = 'GL_Account', how = 'left')
    # Las cuentas que no existen o cuentas nuevas, aparecen con un NaN, se reemplaza el NaN por Nueva
    # si va result = result.fillna('Nueva')
    # Se Filtran las cuentas nuevas, se cambian los nombres de las columnas y se agregan las columnas Country y CoCd
    # si va result = result[result['GL_Account'] == 'Nueva']
    # si va result = result.rename(columns={"GL_Account": "Description"})
    # si va result = result.rename(columns={"Account": "GL_Account"})
    # si va result['Country'] = 'Seleccionar'
    # si va result['CoCd'] = 'Seleccionar'

    #Se despliega el dataframe de "result" en donde se pueden editar las celdas, para que puedan agregar la descripcion, el country y CoCd de cada cuenta nueva
    # si va st.subheader('Cuentas Nuevas')
    # si va st.write("Clasificar las Cuentas Nuevas con el Company Code y el Pais al que corresponden")
    # si va result = st.data_editor(result)
    #
    # si va Company_codes = df_FBL3N[['Company Code']].drop_duplicates()
    # si va groupby_column = st.selectbox('Selecciona la CoCode', Company_codes)
    
    

    # si va df_parametros = pd.concat([df_parametros, result])

    # si va st.dataframe(df_parametros)
    # si va st.write(df_parametros.shape)

    #Nueva columna con la clave = Company Code & Document Number
    # si va FBL3N_merged = df_FBL3N.merge(df_parametros, left_on='Account', right_on='GL_Account', how='left')
    # si va FBL3N_merged['Key'] = FBL3N_merged['Company Code'] + FBL3N_merged['Document Number']
    # si va FBL3N_merged = FBL3N_merged.rename(columns={"CoCd": "Related Party"})
    # si va st.dataframe(FBL3N_merged)

    # si va FBL3N_merged = FBL3N_merged[FBL3N_merged['Company Code'] == groupby_column]
    # si va FBL3N_merged = FBL3N_merged.groupby(by=['Company Code', 'Related Party'], as_index=False)['Amount in local currency'].sum()
    # si va st.dataframe(FBL3N_merged)

    #codigo para editar el dataframe result considerando que hay que agregar la descripcion, en el Country y Cocode que sea un selectbox
    #result = st.data_editor(result)
    #result,
    #column_config={
    #    "Description": st.column_config.TextColumn("Description", help="Copia y pega de SAP la descripcion de la cuenta",),
    #    "Country": st.column_config.SelectboxColumn("Country", help="Selecciona el pais de lista", options=[Company_codes],),
    #    },
    #    disabled=["GL_Account"],
    #    hide_index=True,
    #    )

  
    #new_parametros = st.data_editor(df_parametros)
    #,
    #column_config={
    #    "GL_Account": "GL_Account",
    #    "Description": st.text_input(
    #        "Descrption",
    #        help="Copia y pega la descripcion de la cuenta desde SAP",
    #        width="medium",
    #        ),
    #    "is_widget": "Widget ?",
    #},
    #disabled=["command", "is_widget"],
    #hide_index=True,
    #)







    
    #FBL3N_ctas = df_FBL3N['Account'].astype(str)
    #Parametros_ctas = df_parametros['GL_Account'].astype(str)
    #Ctas_unicas = pd.unique(FBL3N_ctas[['Account']].values.ravel())
    #result = pd.merge(Ctas_unicas, Parametros_ctas, left_on = 'Account', right_on = 'GL_Account', how = 'left')
    #st.dataframe('result')




    # si va edited_df = st.data_editor(
    # si va df_FBL3N,
    # si va column_config={
    # si va     "Company Code": "CoCode",
    # si va     "Document Number": st.column_config.SelectboxColumn(
    # si va         "Doc Number",
    # si va         help="Clasifica el Num de documento",
    # si va         width="medium",
    # si va         options=[
    # si va             "Venta",
    # si va             "Compra",
    # si va             "Hedge",
    # si va         ],
    # si va     ),
    # si va     "is_widget": "Widget ?",
    # si va },
    # si va disabled=["command", "is_widget"],
    # si va hide_index=True,
    # si va )
    
    #groupby_column = st.selectbox(
    #    'What would you like to analyse?',
    #    ('Company Code', 'Account', 'User Name', 'Tax Code'),
    #)

    
    
    
    # -- GROUP DATAFRAME
    #output_columns = ['Amount in local currency']
    #df_grouped_FBL3N = df_FBL3N.groupby(by=[groupby_column], as_index=False)[output_columns].sum()
    ##st.dataframe(df_grouped_FBL3N)

    # -- Información filtrada por company code y agrupada
    #df2 = pd.unique(df_FBL3N[['Company Code']].values.ravel())
    ##st.dataframe(df2)
    
    
    
    ##cocode = st.selectbox('Company Code',df2)
    #cocode = df_FBL3N['Company Code'] == st.selectbox('Choose all Company Codes', df2)
        
    #st.subheader('Auxiliar FBL3N Filtrado por Company code')
    #df_FBL_filtered = df_FBL3N[[cocode]]
    #st.dataframe(df_FBL_filtered)
    
    
    #st.subheader('Gráfica')
    ## -- PLOT DATAFRAME
    #fig = px.bar(
    #    df_grouped_FBL3N,
    #    x=groupby_column,
    #    y='Amount in local currency',
    #    color='Amount in local currency',
    #    color_continuous_scale=['purple', 'green'],
    #    template='plotly_white',
    #    title=f'<b>Sales & Profit by {groupby_column}</b>'
    #)
    #st.plotly_chart(fig)

    ## -- DOWNLOAD SECTION
    
    #def generate_excel_download_link(df_grouped_FBL3N):
        # Credit Excel: https://discuss.streamlit.io/t/how-to-add-a-download-excel-csv-function-to-a-button/4474/5
    #    towrite = BytesIO()
    #    df_grouped_FBL3N.to_excel(towrite, index=False, header=True)  # write to BytesIO buffer
    #    towrite.seek(0)  # reset pointer
    #    b64 = base64.b64encode(towrite.read()).decode()
    #    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Download Excel File</a>'
    #    return st.markdown(href, unsafe_allow_html=True)

    
    
    #st.subheader('Downloads:')
    #generate_excel_download_link(df_grouped_FBL3N)
    ##generate_html_download_link(fig)
