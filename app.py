import streamlit as st
import pandas as pd
from sqlalchemy import create_engine
import plotly.express as px
from datetime import datetime, timedelta
import io
import xlsxwriter

# Configuração da Página
st.set_page_config(page_title="Monisat - Controle de Conversas Atrasadas", layout="wide", page_icon="📊")

# --- FUNÇÃO GERADORA DE EXCEL FORMATADO ---
def gerar_excel_formatado(df_dados):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Relatório Turnos')

    # --- Estilos (Formatação) ---
    fmt_titulo = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter'})
    fmt_subtitulo = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': '#003366', 'bottom': 2})
    fmt_header = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#003366', 'align': 'center', 'border': 1})
    fmt_texto = workbook.add_format({'border': 1, 'align': 'left'})
    fmt_numero = workbook.add_format({'border': 1, 'align': 'center'})
    fmt_total = workbook.add_format({'bold': True, 'bg_color': '#E0E0E0', 'border': 1, 'align': 'center'})

    # --- Cabeçalho do Relatório ---
    # Tenta inserir logo (se não existir, segue sem erro)
    try:
        worksheet.insert_image('A1', 'logo.png', {'x_scale': 0.5, 'y_scale': 0.5})
        worksheet.set_row(0, 50) # Altura da linha do logo
    except:
        pass

    worksheet.merge_range('B2:E2', 'RELATÓRIO DE CONVERSAS ATRASADAS - MONISAT', fmt_titulo)
    worksheet.merge_range('B3:E3', f'Gerado em: {datetime.now().strftime("%d/%m/%Y %H:%M")}', workbook.add_format({'align': 'center'}))

    linha_atual = 5

    # --- Loop por Turno ---
    turnos = ['Manhã', 'Tarde', 'Madrugada']
    
    for turno in turnos:
        # Filtra dados do turno
        df_turno = df_dados[df_dados['turno'] == turno].copy()
        
        if not df_turno.empty:
            # Agrupa por atendente para somar os atrasos
            resumo = df_turno.groupby('atendente')['msg_atrasadas'].sum().reset_index().sort_values(by='msg_atrasadas', ascending=False)
            total_turno = resumo['msg_atrasadas'].sum()

            # Escreve Título do Turno
            worksheet.write(linha_atual, 1, f"TURNO: {turno.upper()}", fmt_subtitulo)
            linha_atual += 2

            # Cabeçalhos da Tabela
            worksheet.write(linha_atual, 1, "Atendente", fmt_header)
            worksheet.write(linha_atual, 2, "Total Atrasos", fmt_header)
            linha_atual += 1

            # Dados
            for _, row in resumo.iterrows():
                worksheet.write(linha_atual, 1, row['atendente'], fmt_texto)
                worksheet.write(linha_atual, 2, row['msg_atrasadas'], fmt_numero)
                linha_atual += 1

            # Total do Turno
            worksheet.write(linha_atual, 1, "TOTAL DO TURNO", fmt_total)
            worksheet.write(linha_atual, 2, total_turno, fmt_total)
            linha_atual += 3 # Espaço para o próximo turno
        
    # Ajustar largura das colunas
    worksheet.set_column('B:B', 30)
    worksheet.set_column('C:C', 15)

    workbook.close()
    return output.getvalue()

# --- CARREGAMENTO DE DADOS ---
@st.cache_data(ttl=60)
def carregar_dados():
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

# --- INTERFACE VISUAL ---
st.title("📊 Monitoramento de Atrasos - Monisat")

if st.button("🔄 Atualizar Dados Agora"):
    st.cache_data.clear()
    st.rerun()

df = carregar_dados()

if df.empty:
    st.warning("Aguardando dados do Robô...")
else:
    # --- FILTROS ---
    st.sidebar.header("Filtros do Relatório")
    periodo = st.sidebar.selectbox("Selecionar Período", ["Hoje", "Ontem", "Mês Atual", "Todo o Histórico"])
    
    # Lógica de Filtro de Data
    hoje = pd.Timestamp.now(tz='America/Sao_Paulo').normalize().tz_localize(None)
    df['data_hora'] = df['data_hora'].dt.tz_localize(None) # Remove fuso para comparar
    
    if periodo == "Hoje":
        df_filtrado = df[df['data_hora'].dt.date == hoje.date()]
    elif periodo == "Ontem":
        ontem = hoje - timedelta(days=1)
        df_filtrado = df[df['data_hora'].dt.date == ontem.date()]
    elif periodo == "Mês Atual":
        df_filtrado = df[(df['data_hora'].dt.month == hoje.month) & (df['data_hora'].dt.year == hoje.year)]
    else:
        df_filtrado = df

    # --- DASHBOARD ---
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Total Atrasos", df_filtrado['msg_atrasadas'].sum())
    kpi2.metric("Ocorrências", len(df_filtrado))
    
    if not df_filtrado.empty:
        pior_atendente = df_filtrado.groupby('atendente')['msg_atrasadas'].sum().idxmax()
        kpi3.metric("Maior Ofensor", pior_atendente)

    st.markdown("---")
    
    # --- ÁREA DE DOWNLOAD DO RELATÓRIO ---
    st.subheader("📂 Exportar Relatório Executivo")
    col_download, col_info = st.columns([1, 3])
    
    with col_download:
        # Gera o arquivo Excel na memória quando o botão é renderizado
        excel_data = gerar_excel_formatado(df_filtrado)
        
        st.download_button(
            label="📥 Baixar Planilha Formatada (.xlsx)",
            data=excel_data,
            file_name=f"Relatorio_Monisat_{periodo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    
    with col_info:
        st.info("Este botão gera um arquivo Excel formatado com o logo da Monisat, separado por turnos e totalizado, pronto para envio.")

    st.markdown("---")

    # --- VISUALIZAÇÃO GRÁFICA ---
    col1, col2 = st.columns([2,1])
    with col1:
        st.subheader("Evolução")
        grafico = df_filtrado.set_index('data_hora').resample('H')['msg_atrasadas'].sum().reset_index()
        fig = px.bar(grafico, x='data_hora', y='msg_atrasadas', title="Atrasos por Hora")
        st.plotly_chart(fig, use_container_width=True)
        
    with col2:
        st.subheader("Ranking")
        ranking = df_filtrado.groupby('atendente')['msg_atrasadas'].sum().sort_values(ascending=False).reset_index()
        st.dataframe(ranking, use_container_width=True, hide_index=True)
