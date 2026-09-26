# app.py Versão 15.0 - Refatorado com motor_remocao

import streamlit as st
import base64
import pandas as pd
from html import escape

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
    montar_cadeias,
    padronizar_texto,
    aplicar_congelamento
)

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Análise de Remoção",
    page_icon="logo_trf1.png",
    layout="wide"
)

# --- AJUSTES VISUAIS ---
# Cores base vêm do tema em .streamlit/config.toml (fonte única)
st.markdown(f"""
    <style>
    :root {{
        --color-primary: {st.get_option("theme.primaryColor")};
        --color-text: {st.get_option("theme.textColor")};
        --color-surface: {st.get_option("theme.secondaryBackgroundColor")};
    }}
    </style>
""", unsafe_allow_html=True)
st.markdown("""
    <style>
    :root {
        --type-title: clamp(1.75rem, 2.5vw, 2.25rem);
        --type-section: 1.5rem;
        --type-subsection: 1.25rem;
        --type-body: 1rem;
        --type-label: 0.875rem;
        --type-caption: 0.875rem;
        --leading-heading: 1.25;
        --leading-body: 1.6;
        --space-1: 0.25rem;
        --space-2: 0.5rem;
        --space-3: 0.75rem;
        --space-4: 1rem;
        --space-6: 1.5rem;
        --space-8: 2rem;
        --space-12: 3rem;
        --color-text-muted: #56626F;
        --color-surface: #F3F6F8;
        --color-border: #D3DAE1;
        --color-warning-surface: #FFF2DA;
        --color-warning-text: #6B4500;
        --radius: 0.5rem;
    }
    [data-testid="stMainBlockContainer"] {
        padding-top: var(--space-8);
        padding-bottom: var(--space-12);
        max-width: 90rem;
        margin-inline: auto;
    }
    .st-key-configuracao {
        width: 100%;
        max-width: 60rem;
        margin-inline: auto;
    }
    .st-key-configuracao [data-testid="stElementContainer"] {
        width: 100%;
    }
    .st-key-configuracao [data-testid="stButton"] {
        width: 100% !important; /* o Streamlit fixa a largura da página via estilo inline */
        display: flex;
        justify-content: center;
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
    [data-testid="stMainBlockContainer"] button [data-testid="stMarkdownContainer"] > p,
    [data-testid="stMainBlockContainer"] summary [data-testid="stMarkdownContainer"] > p {
        margin-bottom: 0;
    }
    [data-testid="stMainBlockContainer"] [data-testid="stExpander"] summary p {
        font-weight: 600;
    }
    [data-testid="stMainBlockContainer"] [data-testid="stExpanderDetails"] {
        padding-top: var(--space-4);
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
    [data-testid="stCaptionContainer"],
    [data-testid="stFileUploader"] small {
        color: var(--color-text-muted);
    }
    [data-testid="stAlertContentWarning"] {
        color: var(--color-warning-text);
    }
    [data-testid="stFileUploaderDropzoneInstructions"] svg {
        color: var(--color-text-muted);
    }
    [data-testid="stDecoration"] { display: none; }
    .cabecalho {
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: var(--space-3);
        text-align: center;
        margin: var(--space-4) 0 var(--space-8);
    }
    .cabecalho img { width: 248px; height: auto; margin-bottom: var(--space-2); }
    [data-testid="stMainBlockContainer"] .cabecalho h1 {
        margin: 0;
        padding: 0;
    }
    .cabecalho-intro {
        margin: 0;
        font-size: var(--type-body);
        line-height: var(--leading-body);
    }
    .cabecalho [data-testid="stHeaderActionElements"] { display: none; }
    .cadeia { margin-bottom: var(--space-6); }
    .cadeia-titulo {
        font-size: var(--type-label);
        font-weight: 600;
        color: var(--color-text-muted);
        margin: 0 0 var(--space-2);
    }
    .cadeia ol {
        display: flex;
        flex-wrap: wrap;
        gap: var(--space-2);
        list-style: none;
        margin: 0;
        padding: 0;
    }
    .cadeia li {
        display: flex;
        align-items: center;
        gap: var(--space-2);
        margin: 0;
    }
    .cadeia-seta {
        font-size: var(--type-subsection);
        color: var(--color-text-muted);
    }
    .etapa {
        display: flex;
        flex-direction: column;
        gap: var(--space-1);
        width: 15rem;
        padding: var(--space-3) var(--space-4);
        background: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius);
        color: var(--color-text);
        line-height: var(--leading-body);
    }
    .etapa-edital { border-left: 4px solid var(--color-primary); }
    .etapa-sobra {
        background: var(--color-warning-surface);
        border-style: dashed;
    }
    .etapa-rotulo {
        font-size: var(--type-caption);
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: var(--color-text-muted);
    }
    .etapa-sobra .etapa-rotulo { color: var(--color-warning-text); }
    .etapa-unidade { font-size: var(--type-body); font-weight: 600; }
    .etapa-ocupante { font-size: var(--type-label); }
    </style>
""", unsafe_allow_html=True)

def renderizar_cadeias(cadeias):
    """Monta o HTML das cadeias de remoção (uma linha de cartões por cadeia)."""
    blocos = []
    numero = 0
    for cadeia in cadeias:
        etapas = cadeia['Etapas']
        remocoes = sum(1 for e in etapas if e['Magistrado'])
        if cadeia['Permuta']:
            titulo = f"Permuta · {remocoes} remoções · a última vaga fecha o ciclo"
        else:
            numero += 1
            plural = "remoção" if remocoes == 1 else "remoções"
            titulo = f"Cadeia {numero} · {remocoes} {plural}" if remocoes else f"Cadeia {numero} · vaga não preenchida"
        itens = []
        for i, etapa in enumerate(etapas):
            classes = ["etapa"]
            if etapa['Magistrado']:
                rotulo = "Vaga do edital" if i == 0 and not cadeia['Permuta'] else "Vaga aberta"
                ocupante = f"Assumida por {escape(str(etapa['Magistrado']))} ·&nbsp;{etapa['Opção Nº']}ª&nbsp;opção"
            else:
                rotulo = "Vaga do edital · remanescente" if i == 0 else "Vaga remanescente"
                ocupante = "Não preenchida"
                classes.append("etapa-sobra")
            if i == 0 and not cadeia['Permuta']:
                classes.append("etapa-edital")
            seta = '<span class="cadeia-seta" aria-hidden="true">→</span>' if i else ""
            itens.append(
                f'<li>{seta}<div class="{" ".join(classes)}">'
                f'<span class="etapa-rotulo">{rotulo}</span>'
                f'<span class="etapa-unidade">{escape(str(etapa["Unidade"]))}</span>'
                f'<span class="etapa-ocupante">{ocupante}</span></div></li>'
            )
        blocos.append(f'<section class="cadeia"><p class="cadeia-titulo">{titulo}</p><ol>{"".join(itens)}</ol></section>')
    return "".join(blocos)

# --- CABEÇALHO COM LOGO ---
with open("logo_trf1.png", "rb") as arquivo_logo:
    logo_base64 = base64.b64encode(arquivo_logo.read()).decode()
st.markdown(
    f'<header class="cabecalho"><img src="data:image/png;base64,{logo_base64}" '
    'alt="Justiça Federal – Tribunal Regional Federal da 1ª Região">'
    '<h1>Sistema de Análise de Remoção de Magistrados</h1>'
    '<p class="cabecalho-intro">Envie a relação de inscritos em formato <strong>Word (.docx)</strong>, '
    '<strong>Excel (.xlsx)</strong> ou <strong>JSON (.json)</strong>.</p></header>',
    unsafe_allow_html=True
)

# Etapa de configuração em coluna estreita e centralizada; resultados usam a largura total
area_configuracao = st.container(key="configuracao")
uploaded_file = area_configuracao.file_uploader(
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
        
        qtd_vagas = len(vagas_detectadas)
        area_configuracao.info(f"O sistema detectou {qtd_vagas} {'vaga disputada' if qtd_vagas == 1 else 'vagas disputadas'}.")

        # --- ÁREA DE CONFIGURAÇÃO (AUTOMAÇÃO DO CONGELAMENTO) ---
        with area_configuracao.expander("Configurações do edital e regras", expanded=False):
            col_vagas, col_regras = st.columns([0.6, 0.4])
            
            with col_vagas:
                texto_vagas_finais = st.text_area(
                    "Vagas disponíveis (uma por linha)", 
                    value=texto_padrao, 
                    height=150
                )
            
            with col_regras:
                st.subheader("Regra de congelamento", anchor=False)
                st.info("Apenas o motivo 'Remoção' gera congelamento de 1 ano. Demais motivos são neutros.")
                data_referencia = st.date_input(
                    "Data de referência (data da nova remoção)",
                    format="DD/MM/YYYY",
                    help="Data utilizada para calcular o interstício de 1 ano. Se não informada, considera hoje."
                )
                st.caption("Magistrados com dados incompletos ou motivos neutros não serão congelados.")
        
        processar = area_configuracao.button("Iniciar o processamento", type="primary")
        aviso_resultado = st.empty()
        area_resultado = st.empty()

        if processar:
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
                        colunas_citadas = " e ".join(f"“{col}”" for col in colunas_automacao_faltando)
                        st.warning(f"**Automação limitada:** colunas {colunas_citadas} não encontradas. O congelamento automático será desativado para todos.")
                
                # Tratamento de Datas
                df_bruto['Data de Exercício'] = pd.to_datetime(df_bruto['Data de Exercício'], dayfirst=True, errors='coerce')
                
                if df_bruto['Data de Exercício'].isna().all() and not df_bruto.empty:
                    with area_resultado.container():
                        st.error("Nenhuma data de exercício válida foi encontrada no arquivo.")
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
                            st.success("Análise concluída.")
                            tab1, tab2, tab3 = st.tabs(["Quadro de remoções", "Cadeias de remoção", "Logs de auditoria"])

                            with tab1:
                                st.caption("Resultado do processamento da remoção. Trata-se apenas de simulação, e não de resultado oficial.")
                                if magistrados_congelados_nomes:
                                    lista_congelados = "\n".join(f"- {m}" for m in magistrados_congelados_nomes)
                                    st.warning(
                                        "**Magistrados congelados** (remoção há menos de 1 ano). "
                                        "Participam do concurso, mas depois dos demais:\n\n" + lista_congelados
                                    )

                                if magistrados_dados_incompletos:
                                    with st.expander("Alerta de integridade: dados incompletos", expanded=False):
                                        st.write("Os seguintes magistrados possuem dados faltantes em 'início da lotação' ou 'motivo da lotação' e foram considerados **descongelados** por padrão:")
                                        for m in magistrados_dados_incompletos:
                                            st.markdown(f"- {m}")

                                # Exibe as unidades com o nome original (o motor devolve a chave normalizada)
                                nomes_unidades = {}
                                for vaga in vagas_iniciais:
                                    nomes_unidades.setdefault(padronizar_texto(vaga), vaga)
                                for origem in df_resultado['Origem']:
                                    nomes_unidades.setdefault(padronizar_texto(origem), origem)
                                df_exibicao = df_resultado.assign(Destino=df_resultado['Vaga'].map(lambda v: nomes_unidades.get(v, v)))

                                df_exibicao['Opção Nº'] = df_exibicao['Opção Nº'].map(lambda n: f"{n}ª")
                                st.dataframe(df_exibicao[['Nome', 'Origem', 'Destino', 'Opção Nº']], use_container_width=True, hide_index=True)
                                excel_data = gerar_excel_em_memoria(df_resultado.drop(columns=['Vaga']), sobras)
                                st.download_button("Baixar Excel", excel_data, "Resultado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                                st.subheader("Vagas remanescentes", anchor=False)
                                if sobras:
                                    st.dataframe(pd.DataFrame([nomes_unidades.get(v, v) for v in sobras], columns=["Unidade"]), use_container_width=True, hide_index=True)
                                else:
                                    st.info("Nenhuma vaga remanescente.")
                            with tab2:
                                st.caption("Cada cadeia começa numa vaga do edital. Quem assume uma vaga deixa a própria lotação, que se torna a vaga seguinte da cadeia, até restar uma vaga remanescente.")
                                cadeias = montar_cadeias(df_resultado, vagas_iniciais)
                                st.markdown(renderizar_cadeias(cadeias), unsafe_allow_html=True)
                            with tab3:
                                logs_apresentacao = "\n".join(
                                    linha.removeprefix("\u23f8\ufe0f ").removeprefix("\U0001f504 ").removeprefix("\u2705 ")
                                    for linha in log
                                )
                                st.caption("Registro de cada passo do processamento: quem assumiu cada vaga, em qual opção, e qual lotação ficou aberta em seguida. Também mostra trocas por uma opção melhor e os casos em que um magistrado cedeu a vez.")
                                st.text_area("Logs de auditoria", value=logs_apresentacao, height=400, label_visibility="collapsed")
                    else:
                        with area_resultado.container():
                            st.warning("Nenhuma movimentação gerada.")
