import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image





st.set_page_config(page_title='Upspace Report', page_icon=":bar_chart:", layout='wide')
st.title(":bar_chart: Upspace Dashboard :rocket:")


st.markdown("##")


st.markdown("# Report Reparto")


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


left_column2, right_column2 = st.columns(2)

with left_column2:
    bar_chart = px.bar(df8, x='Servizi', y='Ore lavorate', title= 'Ore lavorate per reparti', color_discrete_sequence=["#0083B8"]*len(df8),
    template='plotly_white')
    st.plotly_chart(bar_chart)

with right_column2:
    pie_chart = px.pie(df9, names='Servizi.1', values='Primo margine', title= 'Primo Margine per reparto', template='plotly_white')
    st.plotly_chart(pie_chart)

st.markdown("---")


left_column1, right_column1 = st.columns(2)

with left_column1:    
    st.subheader('Classifica ore lavorate')
    st.dataframe(df8)
    

with right_column1:
    st.subheader('Classifica margine')
    st.dataframe(df9)
