import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image




st.set_page_config(page_title='Upspace Report', page_icon=":bar_chart:", layout='wide')
st.title(":bar_chart: Upspace Dashboard :rocket:")

st.markdown("##")


st.markdown("# Report Personale")


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

left_column2, right_column2 = st.columns(2)

with left_column2:
    bar_chart = px.bar(df5, x='Personale.1', y='Ultimo Margine', title= 'Ultimo Margine', color_discrete_sequence=["#0083B8"]*len(df5),
    template='plotly_white')
    st.plotly_chart(bar_chart)

with right_column2:
    bar_chart1 = px.bar(df6, x='Personale.2', y='Costo Personale', title= 'Costo Personale', color_discrete_sequence=["#0083B8"]*len(df6),
    template='plotly_white')
    st.plotly_chart(bar_chart1)

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