# -*- coding: utf-8 -*-
"""
Created on Tue May 14 15:13:01 2024

@author: naguiar
"""

import streamlit as st
import pandas as pd
import warnings
import plotly.express as px
import plotly.graph_objects as go

warnings.simplefilter(action='ignore', category=UserWarning)
def main():
    st.set_page_config(page_title = "Cotador Property")
    page_2()

        
    
    
def page_2(): 
    st.title("Upload da Tabela Excel")

    # Adiciona um botão para upload
    uploaded_file = st.file_uploader("Selecione um arquivo Excel", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        arquivo = pd.ExcelFile(uploaded_file, engine='openpyxl')
        
        st.title("Estatísticas do Prêmios Líquido")
        abas = arquivo.sheet_names
        abas_geral = [aba for aba in abas if 'Geral' in aba]
        
        mes = st.selectbox("Selecione o mês", abas_geral, index=None, placeholder = 'Escolha uma opção')
        if mes:
            df = open_conditions(arquivo, mes)
            df['Prêmio Liquido'] = pd.to_numeric(df['Prêmio Liquido'], errors='coerce')
            status_lista = df['Status'].dropna().unique().tolist()
            
            status = st.selectbox("Selecione o Status", status_lista, index=None, placeholder = 'Escolha uma opção')
            ramos = list(set(df['Ramo'].tolist()))
            unidades = df['Unid. Comercial'].dropna().unique().tolist()
            if status:
                arrumado = estatisticas(df, status, ramos, unidades)
                planilha(arrumado)
            
            st.title("Estatísticas dos Analistas")
            on = st.toggle("Ver estatísticas por Funcionário")
            if on:
                funcionarios(df)
                
                
            
    
    
def funcionarios(df):
    agrupado = df.groupby('Analista')
    stats = agrupado.agg(Quantidade=('Prêmio Liquido', 'size'),
                          Prêmio_Liquido_media=('Prêmio Liquido', 'mean'),
                          Prêmio_Liquido_soma=('Prêmio Liquido', 'sum')).reset_index()
    stats['Porcentagem_Quantidade'] = (stats['Quantidade']/stats['Quantidade'].sum())*100
    stats['Porcentagem_Premio'] = (stats['Prêmio_Liquido_soma']/stats['Prêmio_Liquido_soma'].sum())*100  

    rotulo = {'Analista': 'Analistas', 'Prêmio_Liquido_soma': 'Prêmio Líquido Acumulado (R$)', 'Quantidade': 'Quantidade de Cotações'}
    


    fig = grafico_func(stats, x='Analista', y1='Prêmio_Liquido_soma', y2='Quantidade', titulo='Rendimento dos funcionários', rotulo=rotulo)
    st.plotly_chart(fig)

    fig = criar_grafico_pizza(stats, analista_col='Analista', porcentagem_col='Porcentagem_Premio')
    st.plotly_chart(fig)
    
    funcionarios = stats['Analista'].unique()
    ramo = df['Ramo'].unique()
    df = df.sort_values(by='Ramo')
    
    func_selecionado = st.selectbox("Selecione o analista para ver os status das cotações", funcionarios, index=None, placeholder = 'Escolha uma opção')
    if func_selecionado:
        filtro = df[df['Analista'] == func_selecionado]
        contagem = filtro['Status'].value_counts()
        st.write(contagem)
    
    
        
def estatisticas(df, estado, ramos, unidades):
    df1 = df[(df['Status'] == estado)]
    contagem = df1.groupby(["Unid. Comercial", "Ramo", "Status"]).size().reset_index(name='Contagem')
    premio = df1.groupby(["Unid. Comercial", "Ramo", "Status"])['Prêmio Liquido'].sum()
    teste = pd.merge(premio, contagem, on=["Unid. Comercial", "Ramo", "Status"])
    teste = teste.sort_values(by='Status').reset_index(drop=True)
    
    arrumado = pd.DataFrame(columns=['Unid. Comercial', 'Ramo', 'Status', 'Prêmio Liquido', 'Contagem'])
    
    verificar = [item for item in unidades if item not in teste['Unid. Comercial'].unique()]
    if len(verificar) > 0:
        for v in verificar:
            for r in ramos:
                nova_linha = pd.DataFrame({'Unid. Comercial': [v], 'Ramo': [r], 'Status': [estado], 'Prêmio Liquido': [0], 'Contagem': [0]})
                teste = pd.concat([teste, nova_linha], ignore_index=True)
    
    categorias = teste['Unid. Comercial'].unique()
    
    for categoria in categorias:
        filtrado = teste[teste['Unid. Comercial'] == categoria]
        faltando = [item for item in ramos if item not in filtrado['Ramo'].unique()]
        
        if len(faltando):
            for f in faltando:
                nova_linha = pd.DataFrame({'Unid. Comercial': [categoria], 'Ramo': [f], 'Status': [estado], 'Prêmio Liquido': [0], 'Contagem': [0]})
                filtrado = pd.concat([filtrado, nova_linha], ignore_index=True)
            arrumado = pd.concat([arrumado, filtrado], ignore_index=True)
        else:
            arrumado = pd.concat([arrumado, filtrado], ignore_index=True)

    fig = grafico(arrumado)
    st.plotly_chart(fig)
    
    return arrumado
def planilha(df):
    df = df.dropna(subset=['Ramo'])
    df = df.sort_values(by='Ramo')
    unidades = df['Unid. Comercial'].unique()
    ramo = df['Ramo'].unique()

    matriz = []
    for u in unidades:
        filtrado = df[df['Unid. Comercial'] == u]
        matriz.append(filtrado['Prêmio Liquido'].tolist())

    criado =  pd.DataFrame(matriz, columns=ramo, index = unidades)
    criado['Total'] = criado.sum(axis=1)
    
    st.write(criado)

    
    
def grafico(source):
    fig = px.bar(source, 
                 x='Unid. Comercial', 
                 y='Prêmio Liquido', 
                 color='Ramo', 
                 title='Prêmio Liquido por Unidade Comercial e Ramo',
                 labels={'Prêmio Liquido': 'Prêmio Líquido (R$)', 'Unid. Comercial': 'Unidade Comercial', 'Ramo': 'Ramo'})
    
    fig.update_layout(barmode='stack')  # Empilhar as barras
    
    return fig


def grafico_func(source, x, y1, y2, titulo, rotulo):
    # Criar o gráfico de barras
    fig = go.Figure()
    
    # Adicionar barras para a primeira coluna (Prêmio_Liquido_soma) no eixo y esquerdo
    fig.add_trace(go.Bar(
        x=source[x],
        y=source[y1],
        name=y1,
        marker_color='rgb(55, 83, 109)'
    ))
    
    fig.add_trace(go.Bar(
        x=source[x],
        y=source[y2],
        name=y2,
        marker_color='rgb(26, 118, 255)',
        yaxis='y2'
    ))
    
    # Configurar layout
    fig.update_layout(
        title=titulo,
        xaxis=dict(title=rotulo[x]),
        yaxis=dict(title=rotulo[y1], side='left', showgrid=False),
        yaxis2=dict(title=rotulo[y2], overlaying='y', side='right', showgrid=False),
        barmode='group'
    )
    
    return fig

def criar_grafico_pizza(df, analista_col, porcentagem_col):
    fig = px.pie(df, names=analista_col, values=porcentagem_col, title='Porcentagem de Cotação por Analistas')
    return fig
    
def open_conditions(arquivo, sheet):
    excel = pd.ExcelFile(arquivo, engine= 'openpyxl')
    df = excel.parse(sheet)
    return df  

main()