# app.py Versão 13.1 - Análise de Remoção de Magistrados com Lógica de "Congelamento"

import streamlit as st
import pandas as pd
from docx import Document
import networkx as nx
from pyvis.network import Network
from openpyxl.styles import Alignment
import io
import os
import unicodedata
import re
import json
import streamlit.components.v1 as components

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Análise de Remoção",
    page_icon="⚖️",
    layout="wide"
)

# --- ESTILIZAÇÃO CSS (BOTÃO AZUL) ---
st.markdown("""
    <style>
    div.stButton > button[kind="primary"] {
        background-color: #007BFF;
        color: white;
        border: none;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #0056b3;
        color: white;
        border: none;
    }
    div.stButton > button[kind="primary"]:focus {
        background-color: #0056b3;
        color: white;
        border: none;
        box-shadow: none;
    }
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÃO DE HIGIENIZAÇÃO ---

def padronizar_texto(texto):
    """
    Transforma qualquer texto em um padrão LIMPO para comparação perfeita.
    """
    if not isinstance(texto, str): return ""
    if not texto: return ""
    
    texto = str(texto).upper()
    texto = texto.replace('–', '-').replace('—', '-')
    if " - " in texto:
        texto = texto.split(" - ")[0]
    texto = texto.replace("ª", "A").replace("º", "O")
    texto = texto.replace(" DA SJ ", " ").replace(" DA SSJ ", " ")
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = " ".join(texto.split())
    return texto

def smart_match(vaga, opcao):
    """
    Verifica se a vaga está contida na opção usando word boundaries.
    Evita que '7A VARA' seja encontrada dentro de '27A VARA'.
    """
    # Escapa caracteres especiais de regex e adiciona word boundaries
    pattern = r'\b' + re.escape(vaga) + r'\b'
    return bool(re.search(pattern, opcao))

# --- FUNÇÕES DE EXTRAÇÃO ---

def ler_arquivo_word(uploaded_file):
    try:
        doc = Document(uploaded_file)
        if len(doc.tables) == 0: return None
        tabela = doc.tables[0]
        dados = []
        chaves = []
        for i, linha in enumerate(tabela.rows):
            textos = [c.text.strip() for c in linha.cells]
            if i == 0:
                chaves = textos
                continue
            dados.append(dict(zip(chaves, textos)))
        return pd.DataFrame(dados)
    except Exception as e:
        st.error(f"Erro ao ler arquivo Word: {e}")
        return None

def ler_arquivo_excel(uploaded_file):
    try:
        # REVERTIDO PARA COMPATIBILIDADE COM V11: Deixa o Pandas detectar tipos (int/str)
        # Isso garante que a ordenação de matrícula numérica seja igual à versão anterior
        df = pd.read_excel(uploaded_file)
        
        # Limpeza apenas em colunas de texto
        df = df.fillna('')
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace(['nan', 'NaN', 'None'], '')
            
        return df
    except Exception as e:
        st.error(f"Erro ao ler arquivo Excel: {e}")
        return None

def ler_arquivo_json(uploaded_file):
    try:
        dados = json.load(uploaded_file)
        # Se for lista de dicionários, o Pandas engole direto
        if isinstance(dados, list):
            return pd.DataFrame(dados)
        # Se estiver encapsulado
        elif isinstance(dados, dict):
            for chk in dados.values():
                if isinstance(chk, list):
                    return pd.DataFrame(chk)
            return pd.DataFrame([dados])
        return None
    except Exception as e:
        st.error(f"Erro ao ler JSON: {e}")
        return None

def detectar_vagas_do_edital(df):
    vagas_detectadas = set()
    colunas_opcoes = [col for col in df.columns if 'Opção' in col]
    for col in colunas_opcoes:
        valores = df[col].astype(str).unique()
        for val in valores:
            if "DISPONÍVEL" in val.upper():
                nome_limpo = val.split(" - ")[0].strip() 
                vagas_detectadas.add(nome_limpo)
    return list(vagas_detectadas)

def encurtar_nome(texto):
    if not texto: return "EXTERNO"
    texto = padronizar_texto(texto)
    texto = texto.replace("SUBSECAO JUDICIARIA DE ", "").replace("SECAO JUDICIARIA DE ", "")
    texto = texto.replace("SECAO JUDICIARIA DO ", "").replace("MINAS GERAIS", "MG")
    texto = texto.replace("VARA UNICA", "V.UNICA").replace("RELATORIA", "REL")
    if len(texto) > 30:
        return texto[:28] + "..."
    return texto

def remove_acentos(texto):
    if not isinstance(texto, str): return str(texto)
    return ''.join(c for c in unicodedata.normalize('NFD', texto) 
                  if unicodedata.category(c) != 'Mn').lower()

def normalizar_colunas(df):
    """
    Padroniza os nomes das colunas, ignorando case e acentos.
    Ex: 'Início da Lotação' -> 'início da lotação'
    """
    # Mapa: Nome Final (Interno) -> [Lista de Variantes Possíveis]
    # Obs: As variantes serão comparadas sem acento também.
    mapa_desejado = {
        'Nome': ['nome', 'candidato', 'magistrado', 'nome_completo'],
        'Matrícula': ['matrícula', 'matricula'],
        'Lotação Atual': ['lotação atual', 'lotacao atual', 'unidade atual', 'lotacao_atual', 'lotacao'],
        'Data de Exercício': ['data de exercício', 'data de exercicio', 'exercício', 'data_exercicio'],
        'início da lotação': ['início da lotação', 'inicio da lotacao', 'data de início', 'início na unidade', 'inicio_lota', 'iniciolota'],
        'motivo da lotação': ['motivo da lotação', 'motivo da lotacao', 'motivo', 'motivo_lota', 'motivolota']
    }
    
    # 1. Cria mapa "Slug -> Coluna Real do Excel"
    colunas_excel_slugs = {}
    for col in df.columns:
        slug = remove_acentos(col)
        colunas_excel_slugs[slug] = col
        
    novo_mapa_rename = {}
    
    for alvo_interno, variantes in mapa_desejado.items():
        # Se a coluna alvo já existe EXATAMENTE, pula
        if alvo_interno in df.columns:
            continue
            
        encontrou = False
        
        # Tenta casar o alvo_interno (slugified) com as colunas do excel
        slug_alvo = remove_acentos(alvo_interno)
        if slug_alvo in colunas_excel_slugs:
            coluna_real = colunas_excel_slugs[slug_alvo]
            novo_mapa_rename[coluna_real] = alvo_interno
            encontrou = True
            
        # Tenta casar variantes
        if not encontrou:
            for var in variantes:
                slug_var = remove_acentos(var)
                if slug_var in colunas_excel_slugs:
                    coluna_real = colunas_excel_slugs[slug_var]
                    novo_mapa_rename[coluna_real] = alvo_interno
                    break
    
    if novo_mapa_rename:
        df.rename(columns=novo_mapa_rename, inplace=True)
        
    return df

# --- MOTOR DE REMOÇÃO (V13 - ANTIGUIDADE + LOOKAHEAD ANTI-BLOQUEIO) ---

def processar_remocao(df, vagas_iniciais_lista):
    """
    Motor de remoção com Antiguidade Soberana + Lookahead Anti-Bloqueio.
    
    Regras:
    1. Processa candidatos em ordem de antiguidade (Data + Matrícula)
    2. Antes de alocar, verifica se pegar essa vaga bloquearia a 1ª opção do sênior
    3. Se houver bloqueio, cede a vez para o júnior que ocupa a vaga desejada
    """
    vagas_abertas = set([padronizar_texto(v) for v in vagas_iniciais_lista])
    remocoes_confirmadas = {}
    candidatos_cederam_vez = set()  # Matrículas que cederam vez neste ciclo
    log_movimentacoes = []
    ciclo = 0
    houve_movimentacao = True
    
    # Criar mapa de lotação -> matrícula para lookup rápido
    def get_ocupante_lotacao(lotacao_norm):
        """Retorna a matrícula de quem ocupa determinada lotação."""
        for _, juiz in df.iterrows():
            if padronizar_texto(juiz['Lotação Atual']) == lotacao_norm:
                mat = juiz['Matrícula']
                # Se já foi removido, não ocupa mais
                if mat not in remocoes_confirmadas:
                    return mat
        return None
    
    def get_primeira_opcao_disponivel(juiz, vagas_set):
        """Retorna o índice e texto da primeira opção disponível para o juiz."""
        opcoes = juiz['Lista_Opcoes']
        for i, opcao_bruta in enumerate(opcoes):
            opcao_norm = padronizar_texto(opcao_bruta)
            for vaga in vagas_set:
                if smart_match(vaga, opcao_norm):
                    return i, vaga, opcao_norm
        return None, None, None
    
    def detectar_bloqueio(senior, vaga_pretendida, indice_pretendido):
        """
        Verifica se o sênior pegando esta vaga bloquearia sua própria 1ª opção.
        
        Retorna True se:
        1. O sênior tem uma opção MELHOR que esta (índice menor)
        2. Essa opção melhor é ocupada por um júnior
        3. Esse júnior quer a vaga_pretendida
        """
        opcoes = senior['Lista_Opcoes']
        
        # Se já é a 1ª opção disponível, não há bloqueio
        if indice_pretendido == 0:
            return False
        
        # Verificar opções melhores (índices menores)
        for i in range(indice_pretendido):
            if i >= len(opcoes):
                break
            opcao_melhor_norm = padronizar_texto(opcoes[i])
            
            # Quem ocupa esta opção melhor?
            ocupante_mat = get_ocupante_lotacao(opcao_melhor_norm)
            
            if ocupante_mat is None:
                continue  # Ninguém ocupa, não é bloqueio
            
            # O ocupante quer a vaga_pretendida?
            for _, juiz in df.iterrows():
                if juiz['Matrícula'] == ocupante_mat:
                    opcoes_ocupante = juiz['Lista_Opcoes']
                    for j, op_bruta in enumerate(opcoes_ocupante):
                        op_norm = padronizar_texto(op_bruta)
                        if smart_match(vaga_pretendida, op_norm):
                            # BLOQUEIO DETECTADO!
                            # O ocupante da opção melhor do sênior quer esta vaga
                            return True
                    break
        
        return False
    
    while houve_movimentacao:
        houve_movimentacao = False
        ciclo += 1
        candidatos_cederam_vez.clear()
        
        for _, juiz in df.iterrows():
            matricula = juiz['Matrícula']
            nome = juiz['Nome']
            lotacao_atual = juiz['Lotação Atual']
            lotacao_atual_norm = padronizar_texto(lotacao_atual)
            opcoes = juiz['Lista_Opcoes']
            
            # Pular se já cedeu vez neste ciclo
            if matricula in candidatos_cederam_vez:
                continue
            
            # Determinar índice atual se já tem vaga
            if matricula in remocoes_confirmadas:
                indice_atual = remocoes_confirmadas[matricula]['Opção Index']
            else:
                indice_atual = 9999
            
            # Buscar primeira opção disponível
            match_encontrado = False
            vaga_escolhida = ""
            novo_indice = 0
            opcao_final_texto = ""
            
            for i, opcao_bruta in enumerate(opcoes):
                if i >= indice_atual:
                    break  # Só considera opções melhores que a atual
                
                opcao_norm = padronizar_texto(opcao_bruta)
                for vaga_aberta in vagas_abertas:
                    if smart_match(vaga_aberta, opcao_norm):
                        # LOOKAHEAD: Verificar se há bloqueio
                        if detectar_bloqueio(juiz, vaga_aberta, i):
                            # Ceder vez - não pegar esta vaga agora
                            candidatos_cederam_vez.add(matricula)
                            log_movimentacoes.append(f"⏸️ CICLO {ciclo}: {nome} cedeu vez para destravar opção melhor")
                        else:
                            match_encontrado = True
                            vaga_escolhida = vaga_aberta
                            novo_indice = i
                            opcao_final_texto = opcao_norm
                        break
                if match_encontrado or matricula in candidatos_cederam_vez:
                    break
            
            if match_encontrado:
                juiz_ja_tem_vaga = matricula in remocoes_confirmadas
                
                if juiz_ja_tem_vaga:
                    vaga_anterior = remocoes_confirmadas[matricula]['Destino']
                    vagas_abertas.add(vaga_anterior)
                    log_movimentacoes.append(f"🔄 UPGRADE CICLO {ciclo}: {nome} trocou para {opcao_final_texto} (Opção {novo_indice+1})")
                else:
                    log_movimentacoes.append(f"✅ CICLO {ciclo}: {nome} assumiu {opcao_final_texto} (Opção {novo_indice+1})")
                    if lotacao_atual_norm:
                        vagas_abertas.add(lotacao_atual_norm)
                        log_movimentacoes.append(f"   -> Abriu vaga: {lotacao_atual_norm}")
                
                remocoes_confirmadas[matricula] = {
                    'Matrícula': matricula,
                    'Nome': nome,
                    'Origem': lotacao_atual,
                    'Destino': vaga_escolhida,
                    'DestinoTexto': opcao_final_texto,
                    'Opção Index': novo_indice,
                    'Opção Nº': novo_indice + 1,
                    'Ciclo': ciclo
                }
                vagas_abertas.remove(vaga_escolhida)
                houve_movimentacao = True
                break  # Restart
    
    lista_final = []
    for m, dados in remocoes_confirmadas.items():
        lista_final.append({
            'Matrícula': dados['Matrícula'],
            'Nome': dados['Nome'],
            'Origem': dados['Origem'],
            'Destino': dados['DestinoTexto'],
            'Opção Nº': dados['Opção Nº']
        })
    return pd.DataFrame(lista_final), log_movimentacoes, vagas_abertas

# --- EXPORTAÇÃO ---

def gerar_excel_em_memoria(df_resultado, sobras):
    output = io.BytesIO()
    df_sobras = pd.DataFrame(list(sobras), columns=["Vagas que Sobraram"])
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resultado.to_excel(writer, index=False, sheet_name='Resultado')
        df_sobras.to_excel(writer, index=False, sheet_name='Resultado', startcol=8)
        worksheet = writer.sheets['Resultado']
        estilo_esquerda = Alignment(horizontal='left')
        worksheet['I1'].alignment = estilo_esquerda
        for i in range(len(df_sobras)):
            worksheet[f'I{i+2}'].alignment = estilo_esquerda
    return output.getvalue()

def gerar_html_grafo(df_resultado):
    net = Network(height='600px', width='100%', bgcolor='#222222', font_color='white', directed=True)
    net.barnes_hut(gravity=-3000, central_gravity=0.3, spring_length=300)
    for index, row in df_resultado.iterrows():
        partes_nome = row['Nome'].split(" ")
        juiz_curto = partes_nome[0] + " " + partes_nome[-1] 
        origem = encurtar_nome(row['Origem'])
        destino = encurtar_nome(row['Destino'])
        if not origem: origem = "EXTERNO"
        net.add_node(origem, label=origem, color='#ff6b6b', title="Origem") 
        net.add_node(destino, label=destino, color='#51cf66', title="Destino")
        net.add_edge(origem, destino, title=f"Magistrado: {row['Nome']}", label=juiz_curto, color='white',
            font={'size': 9, 'align': 'middle', 'color': 'white', 'background': '#222222', 'strokeWidth': 0})
    caminho_temp = "grafo_temp.html"
    net.save_graph(caminho_temp)
    with open(caminho_temp, 'r', encoding='utf-8') as f:
        html_string = f.read()
    os.remove(caminho_temp)
    return html_string

# --- CABEÇALHO COM LOGO ---
col_logo, col_titulo = st.columns([1, 3])
with col_logo:
    st.image("logo_trf1.png", width=200)
with col_titulo:
    st.write("")  # Espaço para alinhar com o logo
    st.markdown('<h1 style="color: #002F6C;">Sistema de Análise de Remoção de Magistrados</h1>', unsafe_allow_html=True)

st.markdown("Faça upload do arquivo de inscritos (**Word .docx**, **Excel .xlsx** ou **JSON .json**).")

uploaded_file = st.file_uploader("Arraste o arquivo aqui", type=["docx", "xlsx", "json"])

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
        
        st.info(f"🔎 O sistema detectou {len(vagas_detectadas)} vagas sendo disputadas.")
        
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
                    help="Data utilizada para calcular o interstício de 1 ano. Se não informada, considera hoje."
                )
                st.caption("ℹ️ Magistrados com dados incompletos ou motivos neutros não serão congelados.")
        
        if st.button("🚀 Iniciar Processamento da Remoção", type="primary"):
            with st.spinner('Processando...'):
                vagas_iniciais = [v.strip() for v in texto_vagas_finais.split('\n') if v.strip()]
                
                # --- VALIDAÇÃO DE COLUNAS OBRIGATÓRIAS ---
                colunas_obrigatorias = ['Nome', 'Matrícula', 'Lotação Atual', 'Data de Exercício']
                colunas_automacao = ['início da lotação', 'motivo da lotação']
                
                colunas_faltando = [col for col in colunas_obrigatorias if col not in df_bruto.columns]
                colunas_automacao_faltando = [col for col in colunas_automacao if col not in df_bruto.columns]
                
                if colunas_faltando:
                    st.error(f"❌ **Erro no arquivo:** As seguintes colunas obrigatórias não foram encontradas:")
                    for col in colunas_faltando:
                        st.markdown(f"- `{col}`")
                    st.info("📋 **Colunas encontradas no arquivo:** " + ", ".join(df_bruto.columns.tolist()))
                    st.stop()
                
                if colunas_automacao_faltando:
                    st.warning(f"⚠️ **Automação Limitada:** Colunas `{', '.join(colunas_automacao_faltando)}` não encontradas. O congelamento automático será desativado para todos.")
                
                # Tratamento de Datas
                df_bruto['Data de Exercício'] = pd.to_datetime(df_bruto['Data de Exercício'], dayfirst=True, errors='coerce')
                
                if df_bruto['Data de Exercício'].isna().all() and not df_bruto.empty:
                    st.error("ERRO: Datas inválidas.")
                else:
                    # --- APLICAÇÃO DA LÓGICA DE CONGELAMENTO AUTOMÁTICO ---
                    df_bruto['Status_Congelado'] = 0
                    magistrados_congelados_nomes = []
                    magistrados_dados_incompletos = []
                    
                    ref_date = pd.to_datetime(data_referencia)
                    
                    for idx, row in df_bruto.iterrows():
                        motivo = str(row.get('motivo da lotação', ''))
                        motivo_norm = remove_acentos(motivo)
                        data_inicio = pd.to_datetime(row.get('início da lotação', None), dayfirst=True, errors='coerce')
                        
                        # In Dubio Pro Candidato: Dados faltantes ou inválidos -> Descongelado
                        if pd.isna(data_inicio) or not motivo_norm or motivo_norm == 'nan':
                            if 'início da lotação' in df_bruto.columns or 'motivo da lotação' in df_bruto.columns:
                                magistrados_dados_incompletos.append(row['Nome'])
                            continue
                        
                        # Regra: Se contiver 'remocao' (com ou sem acento), congela.
                        if 'remocao' in motivo_norm:
                            diferenca_dias = (ref_date - data_inicio).days
                            if diferenca_dias < 365:
                                df_bruto.at[idx, 'Status_Congelado'] = 1
                                magistrados_congelados_nomes.append(f"{row['Nome']} ({diferenca_dias} dias)")

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
                        st.success("✅ Análise concluída!")
                        tab1, tab2, tab3 = st.tabs(["📋 Quadro de Remoções", "📊 Resultado Visual", "📜 Logs"])
                        
                        with tab1:
                            if magistrados_congelados_nomes:
                                with st.warning("❄️ Magistrados Congelados Automáticos (1 ano):"):
                                    for m in magistrados_congelados_nomes:
                                        st.markdown(f"- {m}")
                            
                            if magistrados_dados_incompletos:
                                with st.expander("⚠️ Alerta de Integridade: Dados Incompletos", expanded=False):
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
                                st.download_button("📥 Baixar Excel", excel_data, "Resultado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        with tab2:
                            html_grafo = gerar_html_grafo(df_resultado)
                            components.html(html_grafo, height=650, scrolling=True)
                            st.download_button("📥 Grafo (HTML)", html_grafo, "Grafo.html", "text/html")
                        with tab3:
                            st.text_area("Logs", value="\n".join(log), height=400)
                    else:
                        st.warning("Nenhuma movimentação gerada.")