import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image


st.set_page_config(page_title='Upspace Report', page_icon=":bar_chart:", layout='wide')
st.title(":bar_chart: Upspace Dashboard :rocket:")

st.markdown("##")
st.header("Seleziona il report dal menù a tendina :open_file_folder:")



with st.expander("Report Servizi"):
       st.markdown("# Report Servizi :chart_with_upwards_trend:")
       
       st.markdown('##')
       st.markdown('##')
       
       ### --- LOAD DATAFRAME
       excel_file = 'Upspace.xlsx'
       sheet_name = 'Servizi'
       df = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='L:M', header=1, nrows=12)
       df1 = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='P:Q', header=1, nrows=12)
       df2 = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='T:U', header=1, nrows=12)
       df3 = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='E:I', header=0, nrows=38)
       
       ### --- KPI
       sales = int(df3["Vendita"].sum())
       ore = int(df3["Ore lavorate"].sum())
       servizi = int(df3["Quantità"].sum())
       left_column, middle_column, right_column = st.columns(3)
       with left_column:
                 st.subheader('Totale servizi in Euro')
                 st.header(f'€ {sales:,}')
                 
       with middle_column:
                   st.subheader('Totale ore lavorate')
                   st.header(f' {ore:,}')
                   
       with right_column:
            st.subheader('Quantità servizi venduti')
            st.header(servizi)
            
       st.markdown("---")
            
       
       bar_chart = px.bar(df, x='Quantità', y='Servizio', title= 'Classsifica Servizi per quantità venduta', orientation="h", color_discrete_sequence=["#0083B8"]*len(df),
        template='plotly_white')
       
       
    
       
       bar_chart1 = px.bar(df1, x='Margine', y='Servizio.1', title= 'Classsifica Margine per servizi', orientation="h", color_discrete_sequence=["#0083B8"]*len(df),
        template='plotly_white')
       
       left_column2, right_column2 = st.columns(2)    
       left_column2.plotly_chart(bar_chart, use_container_width=True)
       right_column2.plotly_chart(bar_chart1, use_container_width=True)
       
       
       st.markdown("---")
       
       left_column1, middle_column1, right_column1 = st.columns(3)

       with left_column1:    
        st.subheader('Classifica servizi per quantità venduta')
        st.dataframe(df)


       with middle_column1:
        st.subheader('Classifica servizi in ordine di margine')
        st.dataframe(df1)
    


       with right_column1:
        st.subheader('Classifica ore lavorate per ogni servizio')
        st.dataframe(df2)

with st.expander("Report Personale"):
      
      st.markdown("# Report Personale :chart_with_upwards_trend:")


      st.markdown('##')
      st.markdown('##')

      excel_file = 'Upspace.xlsx'
      sheet_name = 'Personale'
      df4 = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='K:L', header=1, nrows=12)
      df5 = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='N:O', header=1, nrows=6)
      df6 = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='Q:R', header=1, nrows=6)
      df7 = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='E:H', header=0, nrows=12)
      
      ore = int(df7["Ore lavorate"].sum())
      personale = int(df7["Costo personale"].sum())
      margine = int(df7["Ultimo margine"].sum())
      left_column, middle_column, right_column = st.columns(3)
      
      with left_column:
                 st.subheader('Totale ore')
                 st.header(f' {ore:,}')
               
      with middle_column:
                   st.subheader('Totale costo personale')
                   st.header(f'€ {personale:,}')


      with right_column:
          st.subheader('Totale margine ultimo')
          st.header(f'€ {margine:,}')
          
      st.markdown("---")
          
      
          
     
      bar_chart = px.bar(df5, x='Personale.1', y='Ultimo Margine', title= 'Ultimo Margine', color_discrete_sequence=["#0083B8"]*len(df5),
               template='plotly_white')
      
     
      bar_chart1 = px.bar(df6, x='Personale.2', y='Costo Personale', title= 'Costo Personale', color_discrete_sequence=["#0083B8"]*len(df6),
               template='plotly_white')
      

      left_column2, right_column2 = st.columns(2)

      left_column2.plotly_chart(bar_chart, use_container_width=True)
      right_column2.plotly_chart(bar_chart1, use_container_width=True)
      
      
      
      
      st.markdown("---")



      left_column1, middle_column1, right_column1 = st.columns(3)

      with left_column1:    
          st.subheader('Ore lavorate per persona')
          st.dataframe(df4)


      with middle_column1:
          st.subheader('Margine per persona')
          st.dataframe(df5)
    


      with right_column1:
          st.subheader('Costo personale')
          st.dataframe(df6)


with st.expander("Report Reparti"):
      
      st.markdown("# Report Reparti :chart_with_upwards_trend:")

      st.markdown('##')
      st.markdown('##')
      
      excel_file = 'Upspace.xlsx'
      sheet_name = 'Reparto'

      df8 = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='L:M', header=1, nrows=6)
      df9 = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='P:Q', header=1, nrows=6)
      df10 = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='F:H', header=0, nrows=7)

      ore = int(df10["Ore lavorate"].sum())
      margine = int(df10["Primo margine"].sum())

      left_column, right_column = st.columns(2)
      
      with left_column:
                 st.subheader('Totale ore')
                 st.header(f' {ore:,}')
                 


      with right_column:
          st.subheader('Totale margine ultimo')
          st.header(f'€ {margine:,}')


      st.markdown("---")
      
      

      
      bar_chart = px.bar(df8, x='Servizi', y='Ore lavorate', title= 'Ore lavorate per reparti', color_discrete_sequence=["#0083B8"]*len(df8),
        template='plotly_white')
     

      
      pie_chart = px.pie(df9, names='Servizi.1', values='Primo margine', title= 'Primo Margine per reparto', template='plotly_white')
     

      left_column3, right_column3 = st.columns(2)

      left_column3.plotly_chart(bar_chart, use_container_width=True)
      right_column3.plotly_chart(pie_chart, use_container_width=True)

     
     
      st.markdown("---")


      left_column1, right_column1 = st.columns(2)

      with left_column1:    
          st.subheader('Classifica ore lavorate')
          st.dataframe(df8)
    
      with right_column1:
          st.subheader('Classifica margine')
          st.dataframe(df9)