# app.py Versão 15.0 - Refatorado com motor_remocao

import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

# Importando o motor "Cérebro"
from motor_remocao import (
    ler_arquivo_word, 
    ler_arquivo_excel, 
    ler_arquivo_json,
    normalizar_colunas, 
    detectar_vagas_do_edital, 
    remove_acentos,
    processar_remocao, 
    gerar_excel_em_memoria, 
    gerar_html_grafo,
    aplicar_congelamento
)

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Análise de Remoção",
    page_icon="logo_trf1.png",
    layout="wide"
)

# --- AJUSTES VISUAIS ---
st.markdown("""
    <style>
    :root {
        --type-title: clamp(1.75rem, 2.5vw, 2.25rem);
        --type-section: 1.5rem;
        --type-subsection: 1.25rem;
        --type-body: 1rem;
        --type-label: 0.875rem;
        --type-caption: 0.8125rem;
        --leading-heading: 1.25;
        --leading-body: 1.5;
        --space-1: 0.25rem;
        --space-2: 0.5rem;
        --space-3: 0.75rem;
        --space-4: 1rem;
        --space-6: 1.5rem;
        --space-8: 2rem;
        --space-12: 3rem;
    }
    [data-testid="stMainBlockContainer"] {
        padding-top: var(--space-8);
        padding-bottom: var(--space-12);
    }
    [data-testid="stMainBlockContainer"] h1,
    [data-testid="stMainBlockContainer"] h2,
    [data-testid="stMainBlockContainer"] h3 {
        font-weight: 600;
        line-height: var(--leading-heading);
        letter-spacing: -0.01em;
    }
    [data-testid="stMainBlockContainer"] h1 {
        font-size: var(--type-title);
        font-weight: 700;
        letter-spacing: -0.02em;
        margin: 0 0 var(--space-4);
    }
    [data-testid="stMainBlockContainer"] h2 {
        font-size: var(--type-section);
        margin: var(--space-8) 0 var(--space-3);
    }
    [data-testid="stMainBlockContainer"] h3 {
        font-size: var(--type-subsection);
        margin: var(--space-6) 0 var(--space-2);
    }
    [data-testid="stMainBlockContainer"] [data-testid="stMarkdownContainer"] > p {
        font-size: var(--type-body);
        line-height: var(--leading-body);
        margin-bottom: var(--space-4);
    }
    [data-testid="stMainBlockContainer"] [data-testid="stWidgetLabel"] p {
        font-size: var(--type-label);
        font-weight: 600;
        line-height: var(--leading-heading);
    }
    [data-testid="stMainBlockContainer"] [data-testid="stCaptionContainer"] {
        font-size: var(--type-caption);
        line-height: var(--leading-body);
    }
    [data-testid="stMainBlockContainer"] [data-testid="stAlert"] {
        font-size: var(--type-body);
        line-height: var(--leading-body);
    }
    div[data-testid="stImage"] {
        margin-top: var(--space-4);
    }
    </style>
""", unsafe_allow_html=True)

# --- CABEÇALHO COM LOGO ---
col_logo, col_titulo = st.columns([1, 4])
with col_logo:
    st.image("logo_trf1.png", width=320)
with col_titulo:
    st.title("Sistema de Análise de Remoção de Magistrados")

st.markdown("Envie a relação de inscritos em formato **Word (.docx)**, **Excel (.xlsx)** ou **JSON (.json)**.")

uploaded_file = st.file_uploader(
    "Arquivo de inscritos",
    type=["docx", "xlsx", "json"],
    help="Selecione ou arraste um arquivo nos formatos indicados."
)

if uploaded_file is not None:
    df_bruto = None
    if uploaded_file.name.endswith('.docx'):
        df_bruto = ler_arquivo_word(uploaded_file)
    elif uploaded_file.name.endswith('.xlsx'):
        df_bruto = ler_arquivo_excel(uploaded_file)
    elif uploaded_file.name.endswith('.json'):
        df_bruto = ler_arquivo_json(uploaded_file)
    
    if df_bruto is not None:
        df_bruto = normalizar_colunas(df_bruto)
        vagas_detectadas = detectar_vagas_do_edital(df_bruto)
        texto_padrao = "\n".join(vagas_detectadas) if vagas_detectadas else ""
        
        st.info(f"O sistema detectou {len(vagas_detectadas)} vagas sendo disputadas.")
        
        # --- ÁREA DE CONFIGURAÇÃO (AUTOMAÇÃO DO CONGELAMENTO) ---
        with st.expander("Configurações do Edital e Regras", expanded=False):
            col_vagas, col_regras = st.columns([0.6, 0.4])
            
            with col_vagas:
                texto_vagas_finais = st.text_area(
                    "Vagas disponíveis (uma por linha):", 
                    value=texto_padrao, 
                    height=150
                )
            
            with col_regras:
                st.write("### Regra de Congelamento")
                st.info("Apenas o motivo 'Remoção' gera congelamento de 1 ano. Demais motivos são neutros.")
                data_referencia = st.date_input(
                    "Data de Referência (Data da Nova Remoção):",
                    format="DD/MM/YYYY",
                    help="Data utilizada para calcular o interstício de 1 ano. Se não informada, considera hoje."
                )
                st.caption("Magistrados com dados incompletos ou motivos neutros não serão congelados.")
        
        aviso_resultado = st.empty()
        area_resultado = st.empty()

        if st.button("Iniciar processamento da remoção", type="primary"):
            with st.spinner('Processando...'):
                vagas_iniciais = [v.strip() for v in texto_vagas_finais.split('\n') if v.strip()]
                
                # --- VALIDAÇÃO DE COLUNAS OBRIGATÓRIAS ---
                colunas_obrigatorias = ['Nome', 'Matrícula', 'Lotação Atual', 'Data de Exercício']
                colunas_automacao = ['início da lotação', 'motivo da lotação']
                
                colunas_faltando = [col for col in colunas_obrigatorias if col not in df_bruto.columns]
                colunas_automacao_faltando = [col for col in colunas_automacao if col not in df_bruto.columns]
                
                if colunas_faltando:
                    with area_resultado.container():
                        st.error("**Erro no arquivo:** As seguintes colunas obrigatórias não foram encontradas:")
                        for col in colunas_faltando:
                            st.markdown(f"- `{col}`")
                        st.info("**Colunas encontradas no arquivo:** " + ", ".join(df_bruto.columns.tolist()))
                    st.stop()
                
                if colunas_automacao_faltando:
                    with aviso_resultado.container():
                        st.warning(f"**Automação limitada:** Colunas `{', '.join(colunas_automacao_faltando)}` não encontradas. O congelamento automático será desativado para todos.")
                
                # Tratamento de Datas
                df_bruto['Data de Exercício'] = pd.to_datetime(df_bruto['Data de Exercício'], dayfirst=True, errors='coerce')
                
                if df_bruto['Data de Exercício'].isna().all() and not df_bruto.empty:
                    with area_resultado.container():
                        st.error("ERRO: Datas inválidas.")
                else:
                    # --- APLICAÇÃO DA LÓGICA DE CONGELAMENTO AUTOMÁTICO (VIA MOTOR) ---
                    df_bruto, magistrados_congelados_nomes, magistrados_dados_incompletos = aplicar_congelamento(df_bruto, data_referencia)

                    # --- ORDENAÇÃO DINÂMICA ---
                    # Status_Congelado (0 primeiro, 1 depois) + Antiguidade
                    criterios_ordenacao = ['Status_Congelado', 'Data de Exercício', 'Matrícula']
                    ascendencia = [True, True, True]
                    
                    df_ordenado = df_bruto.sort_values(by=criterios_ordenacao, ascending=ascendencia)
                    
                    # Limpeza Opções
                    opcoes_limpas = []
                    colunas_opcoes = [col for col in df_bruto.columns if 'Opção' in col]
                    for _, row in df_ordenado.iterrows():
                        lista = [row[col] for col in colunas_opcoes if row[col] and str(row[col]).strip() not in ["", "nan"]]
                        opcoes_limpas.append(lista)
                    df_ordenado['Lista_Opcoes'] = opcoes_limpas
                    
                    # Executa
                    df_resultado, log, sobras = processar_remocao(df_ordenado, vagas_iniciais)
                    
                    if not df_resultado.empty:
                        with area_resultado.container():
                            st.success("Análise concluída!")
                            tab1, tab2, tab3 = st.tabs(["Quadro de remoções", "Resultado visual", "Logs"])

                            with tab1:
                                if magistrados_congelados_nomes:
                                    with st.warning("Magistrados congelados automaticamente (1 ano):"):
                                        for m in magistrados_congelados_nomes:
                                            st.markdown(f"- {m}")

                                if magistrados_dados_incompletos:
                                    with st.expander("Alerta de integridade: dados incompletos", expanded=False):
                                        st.write("Os seguintes magistrados possuem dados faltantes em 'início da lotação' ou 'motivo da lotação' e foram considerados **descongelados** por padrão:")
                                        for m in magistrados_dados_incompletos:
                                            st.markdown(f"- {m}")

                                st.dataframe(df_resultado[['Nome', 'Origem', 'Destino', 'Opção Nº']], use_container_width=True)
                                c1, c2 = st.columns(2)
                                with c1:
                                    if sobras:
                                        st.write("### Vagas Remanescentes")
                                        st.dataframe(pd.DataFrame(list(sobras), columns=["Unidade"]), use_container_width=True)
                                    else: st.info("Sem vagas.")
                                with c2:
                                    excel_data = gerar_excel_em_memoria(df_resultado, sobras)
                                    st.download_button("Baixar Excel", excel_data, "Resultado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            with tab2:
                                html_grafo = gerar_html_grafo(df_resultado)
                                components.html(html_grafo, height=600, scrolling=False)
                                st.download_button("Baixar grafo (HTML)", html_grafo, "Grafo.html", "text/html")
                            with tab3:
                                logs_apresentacao = "\n".join(
                                    linha.removeprefix("\u23f8\ufe0f ").removeprefix("\U0001f504 ").removeprefix("\u2705 ")
                                    for linha in log
                                )
                                st.text_area("Logs", value=logs_apresentacao, height=400)
                    else:
                        with area_resultado.container():
                            st.warning("Nenhuma movimentação gerada.")
