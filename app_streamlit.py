import streamlit as st
import pandas as pd
from datetime import datetime
import numpy as np
import os

# Configuração da página
st.set_page_config(
    page_title="Sistema de Propostas",
    page_icon="",
    layout="wide"
)

# Título da aplicação
st.title("Sistema de Análise de Propostas")

# Caminho da planilha (usando caminho relativo)
caminho_planilha = "PROPOSTAS 2025..xlsx"

# Função para carregar os dados
@st.cache_data
def carregar_dados():
    try:
        # Verifica se o arquivo existe
        if not os.path.exists(caminho_planilha):
            st.error(f"Arquivo não encontrado: {caminho_planilha}")
            st.info("Por favor, certifique-se de que o arquivo 'PROPOSTAS 2025..xlsx' está na mesma pasta do aplicativo.")
            return None
            
        # Lê a planilha pulando as primeiras linhas se necessário
        df = pd.read_excel(caminho_planilha, sheet_name='PROPOSTAS JUD 1')
        
        # Exibe as primeiras linhas para debug
        st.write("Primeiras linhas da planilha:", df.head())
        
        # Verifica se a planilha tem dados
        if df.empty:
            st.error("A planilha está vazia")
            return None
            
        # Renomeia as colunas para os nomes corretos
        # Você precisa ajustar estes nomes de acordo com a estrutura real da sua planilha
        mapeamento_colunas = {
            'MAIO': 'NEGOCIADOR',  # Ajuste conforme necessário
            'Unnamed: 1': 'STATUS',  # Ajuste conforme necessário
            'Unnamed: 2': 'DATA DO ENVIO DA PROPOSTA'  # Ajuste conforme necessário
        }
        
        # Renomeia as colunas
        df = df.rename(columns=mapeamento_colunas)
        
        # Verificando se as colunas necessárias existem
        colunas_necessarias = ['NEGOCIADOR', 'STATUS', 'DATA DO ENVIO DA PROPOSTA']
        for coluna in colunas_necessarias:
            if coluna not in df.columns:
                st.error(f"Coluna '{coluna}' não encontrada na planilha")
                st.info("Verifique se a planilha contém todas as colunas necessárias.")
                return None
        
        # Convertendo a coluna de data
        df['DATA DO ENVIO DA PROPOSTA'] = pd.to_datetime(df['DATA DO ENVIO DA PROPOSTA'], errors='coerce')
        
        # Removendo linhas com datas inválidas
        df = df.dropna(subset=['DATA DO ENVIO DA PROPOSTA'])
        
        # Adicionando coluna de mês
        df['Mês'] = df['DATA DO ENVIO DA PROPOSTA'].dt.strftime('%m/%Y')
        
        # Limpando e padronizando a coluna STATUS
        df['STATUS'] = df['STATUS'].fillna('Não Definido')
        df['STATUS'] = df['STATUS'].astype(str).str.strip()
        
        # Limpando e padronizando a coluna NEGOCIADOR
        df['NEGOCIADOR'] = df['NEGOCIADOR'].fillna('Não Definido')
        df['NEGOCIADOR'] = df['NEGOCIADOR'].astype(str).str.strip()
        
        return df
    except Exception as e:
        st.error(f"Erro ao carregar os dados: {str(e)}")
        st.info("Verifique se o arquivo está no formato correto e se todas as colunas necessárias estão presentes.")
        return None

# Função para gerar o relatório
def gerar_relatorio(dados):
    if dados is None or dados.empty:
        return pd.DataFrame()
    
    try:
        # Agrupando por negociador e status
        resultado = dados.groupby(['NEGOCIADOR', 'STATUS']).size().unstack(fill_value=0)
        
        # Padronizando os nomes das colunas
        colunas_renomear = {}
        if 'Aprovada' in resultado.columns:
            colunas_renomear['Aprovada'] = 'Aprovadas'
        if 'Recusada' in resultado.columns:
            colunas_renomear['Recusada'] = 'Negadas'
        
        resultado = resultado.rename(columns=colunas_renomear)
        
        # Garantindo que as colunas existam
        for col in ['Aprovadas', 'Negadas']:
            if col not in resultado.columns:
                resultado[col] = 0
        
        # Calculando totais e taxas
        resultado['Total de Propostas'] = resultado['Aprovadas'] + resultado['Negadas']
        resultado['Taxa de Aprovação (%)'] = np.where(
            resultado['Total de Propostas'] > 0,
            (resultado['Aprovadas'] / resultado['Total de Propostas'] * 100).round(2),
            0
        )
        
        # Ordenando por total de propostas
        resultado = resultado.sort_values('Total de Propostas', ascending=False)
        
        # Removendo a linha 'vazio' se existir
        if 'vazio ' in resultado.index:
            resultado = resultado.drop('vazio ')
        
        return resultado
    except Exception as e:
        st.error(f"Erro ao gerar relatório: {str(e)}")
        return pd.DataFrame()

# Carregando os dados
df = carregar_dados()

if df is not None and not df.empty:
    # Sidebar para filtros
    st.sidebar.header("Filtros")
    
    # Filtro por negociador
    negociadores = ['Todos'] + sorted(df['NEGOCIADOR'].unique().tolist())
    negociador_selecionado = st.sidebar.selectbox("Selecione o Negociador:", negociadores)
    
    # Filtro por mês
    meses = ['Todos'] + sorted(df['Mês'].unique().tolist())
    mes_selecionado = st.sidebar.selectbox("Selecione o Mês:", meses)
    
    # Aplicando filtros
    df_filtrado = df.copy()
    if negociador_selecionado != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['NEGOCIADOR'] == negociador_selecionado]
    if mes_selecionado != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['Mês'] == mes_selecionado]
    
    # Gerando relatório
    relatorio = gerar_relatorio(df_filtrado)
    
    if not relatorio.empty:
        # Exibindo métricas principais
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total de Propostas", int(relatorio['Total de Propostas'].sum()))
        with col2:
            st.metric("Total Aprovadas", int(relatorio['Aprovadas'].sum()))
        with col3:
            st.metric("Taxa Média de Aprovação", f"{relatorio['Taxa de Aprovação (%)'].mean():.2f}%")
        
        # Exibindo o relatório
        st.subheader("Relatório Detalhado")
        st.dataframe(relatorio, use_container_width=True)
        
        # Botão para download do relatório
        if st.button("Baixar Relatório"):
            try:
                relatorio.to_excel("relatorio_propostas.xlsx")
                st.success("Relatório baixado com sucesso!")
            except Exception as e:
                st.error(f"Erro ao salvar o relatório: {str(e)}")
    else:
        st.warning("Não há dados para exibir com os filtros selecionados.")
else:
    st.error("Não foi possível carregar os dados. Verifique se o arquivo de propostas está no local correto e se está acessível.") 