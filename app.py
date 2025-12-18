import streamlit as st
import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime, timedelta
import io
import xlsxwriter

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Monisat - Controle",
    layout="wide",
    page_icon="logo.png"
)

# --- FUNÇÃO GERADORA DE EXCEL ---
def gerar_excel_formatado(df_dados, titulo_relatorio="Relatório Turnos"):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Relatório')

    # Estilos
    fmt_titulo = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter'})
    fmt_subtitulo = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': '#003366', 'bottom': 2})
    fmt_header = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#003366', 'align': 'center', 'border': 1})
    fmt_texto = workbook.add_format({'border': 1, 'align': 'left'})
    fmt_numero = workbook.add_format({'border': 1, 'align': 'center'})
    fmt_total = workbook.add_format({'bold': True, 'bg_color': '#E0E0E0', 'border': 1, 'align': 'center'})

    try:
        worksheet.insert_image('A1', 'logo.png', {'x_scale': 0.5, 'y_scale': 0.5})
        worksheet.set_row(0, 50)
    except:
        pass

    worksheet.merge_range('B2:E2', 'RELATÓRIO DE MONITORAMENTO - MONISAT', fmt_titulo)
    worksheet.merge_range('B3:E3', f'{titulo_relatorio} - Gerado em: {datetime.now().strftime("%d/%m/%Y %H:%M")}', workbook.add_format({'align': 'center'}))

    linha_atual = 5
    turnos = ['Manhã', 'Tarde', 'Madrugada']
    
    for turno in turnos:
        df_turno = df_dados[df_dados['turno'] == turno].copy()
        if not df_turno.empty:
            resumo = df_turno.groupby('atendente')['msg_atrasadas'].count().reset_index().sort_values(by='msg_atrasadas', ascending=False)
            resumo.rename(columns={'msg_atrasadas': 'ocorrencias'}, inplace=True)
            
            total_turno = resumo['ocorrencias'].sum()

            worksheet.write(linha_atual, 1, f"TURNO: {turno.upper()}", fmt_subtitulo)
            linha_atual += 2
            worksheet.write(linha_atual, 1, "Atendente", fmt_header)
            worksheet.write(linha_atual, 2, "Ocorrências", fmt_header)
            linha_atual += 1

            for _, row in resumo.iterrows():
                worksheet.write(linha_atual, 1, row['atendente'], fmt_texto)
                worksheet.write(linha_atual, 2, row['ocorrencias'], fmt_numero)
                linha_atual += 1

            worksheet.write(linha_atual, 1, "TOTAL (FLAGRANTES)", fmt_total)
            worksheet.write(linha_atual, 2, total_turno, fmt_total)
            linha_atual += 3
        
    worksheet.set_column('B:B', 30)
    worksheet.set_column('C:C', 15)
    workbook.close()
    return output.getvalue()

# --- CARREGAMENTO DE DADOS (CORREÇÃO DE FUSO HORÁRIO AQUI) ---
@st.cache_data(ttl=60)
def carregar_dados():
    db_url = st.secrets["DB_URL"]
    try:
        conn = create_engine(db_url)
        df = pd.read_sql("SELECT * FROM registros", conn)
        
        if not df.empty:
            # 1. Converte para datetime
            df['data_hora'] = pd.to_datetime(df['data_hora'])
            
            # 2. CORREÇÃO DO BUG: Assume que está em UTC e converte para Brasília
            # Se o banco salvou como UTC (sem info de fuso), 'tz_localize' diz "Isso é UTC"
            # e 'tz_convert' diz "Transforme para horário do Brasil (-3h)"
            try:
                df['data_hora'] = df['data_hora'].dt.tz_localize('UTC').dt.tz_convert('America/Sao_Paulo')
            except:
                # Caso já tenha vindo com fuso, apenas converte
                df['data_hora'] = df['data_hora'].dt.tz_convert('America/Sao_Paulo')
                
            # 3. Remove a info de fuso para facilitar comparações simples de data
            df['data_hora'] = df['data_hora'].dt.tz_localize(None)
            
        return df
    except Exception as e:
        st.error(f"Erro ao conectar no banco: {e}")
        return pd.DataFrame()

df = carregar_dados()

# --- BARRA LATERAL (LOGO NO TOPO) ---
try:
    # Logo é a primeira coisa da Sidebar
    st.sidebar.image("logo.png", use_column_width=True) 
except:
    pass

st.sidebar.header("Filtros do Relatório")

if df.empty:
    st.warning("Aguardando dados...")
    df_filtrado = pd.DataFrame()
else:
    periodo = st.sidebar.selectbox("Selecionar Período", ["Hoje", "Ontem", "Mês Atual", "Todo o Histórico"])
    
    if st.sidebar.button("🔄 Atualizar Dados"):
        st.cache_data.clear()
        st.rerun()

    hoje = pd.Timestamp.now(tz='America/Sao_Paulo').normalize().tz_localize(None)
    
    if periodo == "Hoje":
        df_filtrado = df[df['data_hora'].dt.date == hoje.date()]
    elif periodo == "Ontem":
        ontem = hoje - timedelta(days=1)
        df_filtrado = df[df['data_hora'].dt.date == ontem.date()]
    elif periodo == "Mês Atual":
        df_filtrado = df[(df['data_hora'].dt.month == hoje.month) & (df['data_hora'].dt.year == hoje.year)]
    else:
        df_filtrado = df

    # --- BIBLIOTECA DE RELATÓRIOS ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("📂 Histórico Mensal")
    df['mes_ano'] = df['data_hora'].dt.strftime('%m/%Y')
    meses_disponiveis = df['mes_ano'].unique()
    mes_selecionado = st.sidebar.selectbox("Baixar Mês:", meses_disponiveis)
    
    if mes_selecionado:
        df_historico = df[df['mes_ano'] == mes_selecionado]
        excel_historico = gerar_excel_formatado(df_historico, titulo_relatorio=f"Relatório - {mes_selecionado}")
        st.sidebar.download_button(
            label=f"📥 Download {mes_selecionado}",
            data=excel_historico,
            file_name=f"Relatorio_{mes_selecionado.replace('/','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# --- PAINEL PRINCIPAL ---
col_titulo, col_botao = st.columns([3, 1])

with col_titulo:
    st.title("📊 Monitoramento Monisat")

with col_botao:
    st.write("") # Espaço para alinhar
    # --- BOTÃO LIMPO (Sem texto azul) ---
    if not df_filtrado.empty:
        excel_data = gerar_excel_formatado(df_filtrado, titulo_relatorio=f"Relatório - {periodo}")
        st.download_button(
            label="📥 Baixar Relatório (.xlsx)",
            data=excel_data,
            file_name=f"Relatorio_Monisat_{periodo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

if not df_filtrado.empty:
    st.markdown("---")
    
    # --- KPIS ---
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Atrasos (Volume)", df_filtrado['msg_atrasadas'].sum())
    col2.metric("Ocorrências (Flagrantes)", len(df_filtrado))
    pior_atendente = df_filtrado.groupby('atendente').size().idxmax()
    col3.metric("Maior Ofensor (Freq.)", pior_atendente)
    
    st.markdown("---")

    # --- RANKINGS (Sem Gráfico de Evolução) ---
    st.subheader("🏆 Rankings por Ocorrência (Frequência de Atrasos)")
    col_m, col_t, col_n, col_g = st.columns(4)

    def mostrar_ranking_ocorrencia(dataframe, turno_nome, coluna_alvo):
        coluna_alvo.markdown(f"#### {turno_nome}")
        if turno_nome == "Geral":
             df_rank = dataframe.groupby('atendente').size().reset_index(name='ocorrências').sort_values(by='ocorrências', ascending=False)
        else:
             df_rank = dataframe[dataframe['turno'] == turno_nome].groupby('atendente').size().reset_index(name='ocorrências').sort_values(by='ocorrências', ascending=False)
        
        if not df_rank.empty:
            coluna_alvo.dataframe(df_rank, hide_index=True, use_container_width=True)
        else:
            coluna_alvo.info("-")

    mostrar_ranking_ocorrencia(df_filtrado, "Manhã", col_m)
    mostrar_ranking_ocorrencia(df_filtrado, "Tarde", col_t)
    mostrar_ranking_ocorrencia(df_filtrado, "Madrugada", col_n)
    mostrar_ranking_ocorrencia(df_filtrado, "Geral", col_g)

else:
    st.info("Nenhum dado encontrado para o período selecionado.")
