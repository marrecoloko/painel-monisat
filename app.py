import streamlit as st
import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime, timedelta
import io
import xlsxwriter

# --- CONFIGURAÇÃO DA PÁGINA (COM LOGO NA ABA) ---
# Altere 'logo.png' caso o nome do seu arquivo seja diferente
st.set_page_config(
    page_title="Monisat - Controle",
    layout="wide",
    page_icon="logo1.png" 
)

# --- FUNÇÃO GERADORA DE EXCEL FORMATADO ---
def gerar_excel_formatado(df_dados):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Relatório Turnos')

    # Estilos
    fmt_titulo = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter'})
    fmt_subtitulo = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': '#003366', 'bottom': 2})
    fmt_header = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#003366', 'align': 'center', 'border': 1})
    fmt_texto = workbook.add_format({'border': 1, 'align': 'left'})
    fmt_numero = workbook.add_format({'border': 1, 'align': 'center'})
    fmt_total = workbook.add_format({'bold': True, 'bg_color': '#E0E0E0', 'border': 1, 'align': 'center'})

    # Inserir Logo no Excel (se existir)
    try:
        worksheet.insert_image('A1', 'logo.png', {'x_scale': 0.5, 'y_scale': 0.5})
        worksheet.set_row(0, 50)
    except:
        pass

    worksheet.merge_range('B2:E2', 'RELATÓRIO DE MONITORAMENTO - MONISAT', fmt_titulo)
    worksheet.merge_range('B3:E3', f'Gerado em: {datetime.now().strftime("%d/%m/%Y %H:%M")}', workbook.add_format({'align': 'center'}))

    linha_atual = 5
    turnos = ['Manhã', 'Tarde', 'Madrugada']
    
    for turno in turnos:
        df_turno = df_dados[df_dados['turno'] == turno].copy()
        if not df_turno.empty:
            resumo = df_turno.groupby('atendente')['msg_atrasadas'].sum().reset_index().sort_values(by='msg_atrasadas', ascending=False)
            total_turno = resumo['msg_atrasadas'].sum()

            worksheet.write(linha_atual, 1, f"TURNO: {turno.upper()}", fmt_subtitulo)
            linha_atual += 2
            worksheet.write(linha_atual, 1, "Atendente", fmt_header)
            worksheet.write(linha_atual, 2, "Total Atrasos", fmt_header)
            linha_atual += 1

            for _, row in resumo.iterrows():
                worksheet.write(linha_atual, 1, row['atendente'], fmt_texto)
                worksheet.write(linha_atual, 2, row['msg_atrasadas'], fmt_numero)
                linha_atual += 1

            worksheet.write(linha_atual, 1, "TOTAL DO TURNO", fmt_total)
            worksheet.write(linha_atual, 2, total_turno, fmt_total)
            linha_atual += 3
        
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

# --- LÓGICA DO FILTRO E DADOS ---
df = carregar_dados()

# --- BARRA LATERAL (LOGO E FILTROS) ---
# Tenta carregar o logo na sidebar
try:
    st.sidebar.image("logo1.png", use_column_width=True)
except:
    # Se não achar o logo, segue a vida sem erro
    pass

st.sidebar.header("Filtros do Relatório")

if df.empty:
    st.warning("Aguardando dados do Robô...")
    df_filtrado = pd.DataFrame()
else:
    periodo = st.sidebar.selectbox("Selecionar Período", ["Hoje", "Ontem", "Mês Atual", "Todo o Histórico"])
    
    if st.sidebar.button("🔄 Atualizar Dados"):
        st.cache_data.clear()
        st.rerun()

    hoje = pd.Timestamp.now(tz='America/Sao_Paulo').normalize().tz_localize(None)
    df['data_hora'] = df['data_hora'].dt.tz_localize(None)
    
    if periodo == "Hoje":
        df_filtrado = df[df['data_hora'].dt.date == hoje.date()]
    elif periodo == "Ontem":
        ontem = hoje - timedelta(days=1)
        df_filtrado = df[df['data_hora'].dt.date == ontem.date()]
    elif periodo == "Mês Atual":
        df_filtrado = df[(df['data_hora'].dt.month == hoje.month) & (df['data_hora'].dt.year == hoje.year)]
    else:
        df_filtrado = df

# --- LAYOUT DO TOPO (TÍTULO E BOTÃO EXPORTAR) ---
col_titulo, col_botao = st.columns([3, 1])

with col_titulo:
    st.title("📊 Monitoramento conversas atrasadas Monisat")

with col_botao:
    st.write("")
    if not df_filtrado.empty:
        excel_data = gerar_excel_formatado(df_filtrado)
        st.download_button(
            label="📥 Baixar Relatório (.xlsx)",
            data=excel_data,
            file_name=f"Relatorio_Monisat_{periodo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

if not df_filtrado.empty:
    # --- KPIS ---
    st.markdown("---")
    col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
    
    # Total Atrasos (Volume de mensagens)
    col_kpi1.metric("Total Atrasos (Volume)", df_filtrado['msg_atrasadas'].sum(), help="Soma total de todas as mensagens atrasadas detectadas no período.")
    
    # Ocorrências (Frequência de incidentes)
    col_kpi2.metric("Ocorrências (Flagrantes)", len(df_filtrado), help="Quantidade de vezes que o robô detectou um atendente com atraso, independente de quantas mensagens eram.")
    
    pior_atendente = df_filtrado.groupby('atendente')['msg_atrasadas'].sum().idxmax()
    col_kpi3.metric("Maior Quantidade de Atraso", pior_atendente)
    st.markdown("---")

    # --- RANKINGS (4 COLUNAS) ---
    st.subheader("Rankings por Turno e Geral")
    col_m, col_t, col_n, col_g = st.columns(4)

    def mostrar_ranking(dataframe, turno_nome, coluna_alvo):
        coluna_alvo.markdown(f"#### {turno_nome}")
        if turno_nome == "Geral":
             df_rank = dataframe.groupby('atendente')['msg_atrasadas'].sum().sort_values(ascending=False).reset_index()
        else:
             df_rank = dataframe[dataframe['turno'] == turno_nome].groupby('atendente')['msg_atrasadas'].sum().sort_values(ascending=False).reset_index()
        
        if not df_rank.empty:
            coluna_alvo.dataframe(df_rank, hide_index=True, use_container_width=True)
        else:
            coluna_alvo.info("Sem dados.")

    mostrar_ranking(df_filtrado, "Manhã", col_m)
    mostrar_ranking(df_filtrado, "Tarde", col_t)
    mostrar_ranking(df_filtrado, "Madrugada", col_n)
    mostrar_ranking(df_filtrado, "Geral", col_g)

else:
    st.info("Nenhum dado encontrado para o período selecionado.")
