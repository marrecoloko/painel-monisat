import streamlit as st
import pandas as pd
from sqlalchemy import create_engine
import plotly.express as px
from datetime import timedelta

# Configuração da Página
st.set_page_config(page_title="Monisat - Controle", layout="wide")

# Função de carregamento (Cacheado para ser rápido)
@st.cache_data(ttl=60) # Atualiza a cada 60 segundos
def carregar_dados():
    # Pega a senha dos segredos do Streamlit (Segurança)
    db_url = st.secrets["DB_URL"]
    try:
        conn = create_engine(db_url)
        df = pd.read_sql("SELECT * FROM registros", conn)
        if not df.empty:
            df['data_hora'] = pd.to_datetime(df['data_hora'])
        return df
    except Exception as e:
        st.error(f"Erro ao conectar no banco: {e}")
        return pd.DataFrame()

st.title("📊 Monitoramento de Conversas Atrasadas - Monisat")

if st.button("🔄 Atualizar Dados Agora"):
    st.cache_data.clear()
    st.rerun()

df = carregar_dados()

if df.empty:
    st.warning("Aguardando dados do Robô...")
else:
    # --- FILTROS ---
    st.sidebar.header("Filtros")
    periodo = st.sidebar.selectbox("Período", ["Hoje", "Ontem", "Últimos 7 dias", "Tudo"])
    
    hoje = pd.Timestamp.now(tz='America/Sao_Paulo').normalize().tz_localize(None)
    # Ajuste fuso horário simples removendo info de time zone para comparação
    df['data_hora'] = df['data_hora'].dt.tz_localize(None)
    
    if periodo == "Hoje":
        df_filtrado = df[df['data_hora'].dt.date == hoje.date()]
    elif periodo == "Ontem":
        ontem = hoje - timedelta(days=1)
        df_filtrado = df[df['data_hora'].dt.date == ontem.date()]
    elif periodo == "Últimos 7 dias":
        df_filtrado = df[df['data_hora'] >= (hoje - timedelta(days=7))]
    else:
        df_filtrado = df

    # --- DADOS ---
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Total Msg Atrasadas", df_filtrado['msg_atrasadas'].sum())
    kpi2.metric("Ocorrências Registradas", len(df_filtrado))
    
    if not df_filtrado.empty:
        pior_atendente = df_filtrado.groupby('atendente')['msg_atrasadas'].sum().idxmax()
        kpi3.metric("Maior Ofensor", pior_atendente)

    col1, col2 = st.columns([2,1])
    
    with col1:
        st.subheader("Evolução no Tempo")
        grafico = df_filtrado.set_index('data_hora').resample('H')['msg_atrasadas'].sum().reset_index()
        fig = px.bar(grafico, x='data_hora', y='msg_atrasadas')
        st.plotly_chart(fig, use_container_width=True)
        
    with col2:
        st.subheader("Ranking por Atendente")
        ranking = df_filtrado.groupby('atendente')['msg_atrasadas'].sum().sort_values(ascending=False)
        st.dataframe(ranking, use_container_width=True)

    # Exportação
    csv = df_filtrado.to_csv(index=False).encode('utf-8')
    st.download_button("📥 Baixar Relatório em Excel/CSV", csv, "relatorio_monisat.csv", "text/csv")
