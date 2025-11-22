import streamlit as st
import pandas as pd
from docx import Document
import networkx as nx
from pyvis.network import Network
from openpyxl.styles import Alignment
import io
import os
import unicodedata # Biblioteca para remover acentos e caracteres especiais
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

# --- FUNÇÃO DE HIGIENIZAÇÃO (O SEGREDO DA COMPATIBILIDADE) ---

def padronizar_texto(texto):
    """
    Transforma qualquer texto em um padrão LIMPO para comparação perfeita.
    Resolve o problema de 22ª vs 22A e traços diferentes.
    """
    if not isinstance(texto, str): return ""
    if not texto: return ""
    
    # 1. Converter para string e Maiúsculo
    texto = str(texto).upper()
    
    # 2. Padronizar traços (Excel costuma usar travessão '–' em vez de hífen '-')
    texto = texto.replace('–', '-').replace('—', '-')
    
    # 3. Remover sufixos de status (ANTES de limpar acentos para garantir o split no traço certo)
    if " - " in texto:
        texto = texto.split(" - ")[0]
    
    # 4. Substituições manuais críticas para Magistratura
    texto = texto.replace("ª", "A").replace("º", "O") # Resolve 22ª vs 22A
    texto = texto.replace(" DA SJ ", " ").replace(" DA SSJ ", " ") # Remove conectivos
    
    # 5. Normalização Unicode (Remove acentos: Ú -> U, Ç -> C, Ã -> A)
    # NFD separa o caractere base do acento, e filtramos apenas os não-acentos
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    
    # 6. Remove espaços duplos e caracteres invisíveis (Non-breaking space \xa0)
    texto = " ".join(texto.split())
    
    return texto

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
        # Lê o Excel. Não forçamos dtype=str globalmente para não quebrar datas,
        # mas trataremos as strings depois.
        df = pd.read_excel(uploaded_file)
        
        # Limpeza de valores nulos
        df = df.fillna('')
        
        # Converte colunas de texto para string e limpa espaços
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace(['nan', 'NaN', 'None'], '')
            
        return df
    except Exception as e:
        st.error(f"Erro ao ler arquivo Excel: {e}")
        return None

def detectar_vagas_do_edital(df):
    """Detecta vagas baseado na string 'DISPONÍVEL'."""
    vagas_detectadas = set()
    colunas_opcoes = [col for col in df.columns if 'Opção' in col]
    
    for col in colunas_opcoes:
        # Pega valores únicos como string
        valores = df[col].astype(str).unique()
        for val in valores:
            if "DISPONÍVEL" in val.upper():
                # Usa a função padronizar_texto para extrair o nome limpo (Ex: 22A VARA MG)
                # Isso garante que o que aparece na caixa de texto já é o nome "chave" para busca
                nome_limpo = val.split(" - ")[0].strip() 
                # Nota: Aqui mantemos o nome original (sem padronizar_texto total) para o usuário ler,
                # mas a padronização total acontece dentro do processar_remocao
                vagas_detectadas.add(nome_limpo)
                
    return list(vagas_detectadas)

def encurtar_nome(texto):
    """Nome curto para visualização no gráfico."""
    if not texto: return "EXTERNO"
    # Usa o texto já padronizado
    texto = padronizar_texto(texto)
    
    # Remoções estéticas
    texto = texto.replace("SUBSECAO JUDICIARIA DE ", "").replace("SECAO JUDICIARIA DE ", "")
    texto = texto.replace("SECAO JUDICIARIA DO ", "").replace("MINAS GERAIS", "MG")
    texto = texto.replace("VARA UNICA", "V.UNICA").replace("RELATORIA", "REL")
    
    if len(texto) > 30:
        return texto[:28] + "..."
    return texto

# --- MOTOR DE REMOÇÃO ---

def processar_remocao(df, vagas_iniciais_lista):
    # O conjunto de vagas abertas agora usa o texto padronizado (sem acentos, sem ª)
    vagas_abertas = set([padronizar_texto(v) for v in vagas_iniciais_lista])
    
    remocoes_confirmadas = {} 
    log_movimentacoes = []
    ciclo = 0
    houve_movimentacao = True
    
    while houve_movimentacao:
        houve_movimentacao = False
        ciclo += 1
        for index, juiz in df.iterrows():
            matricula = juiz['Matrícula']
            if matricula in remocoes_confirmadas: continue
            
            nome = juiz['Nome']
            lotacao_atual = juiz['Lotação Atual']
            # Padroniza a lotação atual para quando ela virar vaga
            lotacao_atual_norm = padronizar_texto(lotacao_atual)
            opcoes = juiz['Lista_Opcoes']
            
            match_encontrado = False
            vaga_escolhida = ""
            opcao_numero = 0
            opcao_final_texto = ""
            
            for i, opcao_bruta in enumerate(opcoes):
                # Padroniza a opção do candidato (remove acentos, ª, etc)
                opcao_norm = padronizar_texto(opcao_bruta)
                
                for vaga_aberta in vagas_abertas:
                    # Comparação blindada: Texto limpo vs Texto limpo
                    # Usamos 'in' para flexibilidade (ex: "22A VARA" in "22A VARA MG")
                    if vaga_aberta in opcao_norm or opcao_norm in vaga_aberta:
                        match_encontrado = True
                        vaga_escolhida = vaga_aberta
                        opcao_numero = i + 1
                        opcao_final_texto = opcao_norm
                        break
                if match_encontrado: break
            
            if match_encontrado:
                remocoes_confirmadas[matricula] = {
                    'Matrícula': matricula,
                    'Nome': nome,
                    'Origem': lotacao_atual,
                    'Destino': opcao_final_texto,
                    'Opção Nº': opcao_numero,
                    'Ciclo': ciclo
                }
                log_movimentacoes.append(f"✅ CICLO {ciclo}: {nome} assumiu {opcao_final_texto}")
                vagas_abertas.remove(vaga_escolhida)
                
                if lotacao_atual_norm:
                    vagas_abertas.add(lotacao_atual_norm)
                    log_movimentacoes.append(f"   -> Abriu vaga: {lotacao_atual_norm}")
                
                houve_movimentacao = True
                break 
    
    return pd.DataFrame(list(remocoes_confirmadas.values())), log_movimentacoes, vagas_abertas

# --- EXPORTAÇÃO E VISUALIZAÇÃO ---

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
    net.barnes_hut(gravity=-3000, central_gravity=0.3, spring_length=200)

    for index, row in df_resultado.iterrows():
        partes_nome = row['Nome'].split(" ")
        juiz_curto = partes_nome[0] + " " + partes_nome[-1] 
        origem = encurtar_nome(row['Origem'])
        destino = encurtar_nome(row['Destino'])
        if not origem: origem = "EXTERNO"
            
        net.add_node(origem, label=origem, color='#ff6b6b', title="Origem") 
        net.add_node(destino, label=destino, color='#51cf66', title="Destino")
        
        net.add_edge(
            origem, destino, 
            title=f"Magistrado: {row['Nome']}", 
            label=juiz_curto, 
            color='white',
            font={'size': 9, 'align': 'middle', 'color': 'white', 'background': '#222222', 'strokeWidth': 0}
        )
    
    caminho_temp = "grafo_temp.html"
    net.save_graph(caminho_temp)
    with open(caminho_temp, 'r', encoding='utf-8') as f:
        html_string = f.read()
    os.remove(caminho_temp)
    return html_string

# --- INTERFACE PRINCIPAL ---

st.title("⚖️ Sistema de Análise de Remoção de Magistrados")
st.markdown("Faça upload do arquivo de inscritos (**Word .docx** ou **Excel .xlsx**).")

uploaded_file = st.file_uploader("Arraste o arquivo aqui", type=["docx", "xlsx"])

if uploaded_file is not None:
    df_bruto = None
    if uploaded_file.name.endswith('.docx'):
        df_bruto = ler_arquivo_word(uploaded_file)
    elif uploaded_file.name.endswith('.xlsx'):
        df_bruto = ler_arquivo_excel(uploaded_file)
    
    if df_bruto is not None:
        # Autodetecção de Vagas
        vagas_detectadas = detectar_vagas_do_edital(df_bruto)
        texto_padrao = "\n".join(vagas_detectadas) if vagas_detectadas else ""
        
        st.info(f"🔎 O sistema detectou {len(vagas_detectadas)} vagas ofertadas no edital.")
        
        with st.expander("Ver ou Editar Vagas Iniciais Detectadas", expanded=True):
            texto_vagas_finais = st.text_area(
                "Vagas consideradas (uma por linha):",
                value=texto_padrao,
                height=100
            )
        
        if st.button("🚀 Iniciar Processamento da Remoção", type="primary"):
            with st.spinner('Processando...'):
                
                vagas_iniciais = [v.strip() for v in texto_vagas_finais.split('\n') if v.strip()]
                
                # --- TRATAMENTO DE DATAS CRÍTICO ---
                # Garante que datas vindas do Excel ou Word sejam tratadas iguais
                df_bruto['Data de Exercício'] = pd.to_datetime(
                    df_bruto['Data de Exercício'], 
                    dayfirst=True, 
                    errors='coerce'
                )
                
                # --- VERIFICAÇÃO DE SEGURANÇA (NOVO) ---
                # Se todas as datas falharem (forem NaT) e o dataframe não estiver vazio, para tudo.
                if df_bruto['Data de Exercício'].isna().all() and not df_bruto.empty:
                    st.error("ERRO: Não foi possível ler as datas. Verifique se a coluna 'Data de Exercício' está no formato DD/MM/AAAA.")
                else:
                    # Ordenação
                    df_ordenado = df_bruto.sort_values(by=['Data de Exercício', 'Matrícula'], ascending=[True, True])
                    
                    # Limpeza das Opções
                    opcoes_limpas = []
                    colunas_opcoes = [col for col in df_bruto.columns if 'Opção' in col]
                    for _, row in df_ordenado.iterrows():
                        lista = [row[col] for col in colunas_opcoes if row[col] and str(row[col]).strip() not in ["", "nan"]]
                        opcoes_limpas.append(lista)
                    df_ordenado['Lista_Opcoes'] = opcoes_limpas
                    
                    # Executa a Remoção
                    df_resultado, log, sobras = processar_remocao(df_ordenado, vagas_iniciais)
                    
                    if not df_resultado.empty:
                        st.success("✅ Análise concluída!")
                        
                        tab1, tab2, tab3 = st.tabs(["📊 Resultado Visual", "📋 Tabela Oficial", "📜 Logs Detalhados"])
                        
                        with tab1:
                            html_grafo = gerar_html_grafo(df_resultado)
                            components.html(html_grafo, height=650, scrolling=True)
                            st.download_button("📥 Baixar Grafo (HTML)", html_grafo, "Grafo_Remocao.html", "text/html")
                        
                        with tab2:
                            st.dataframe(df_resultado[['Nome', 'Origem', 'Destino', 'Opção Nº']], use_container_width=True)
                            
                            c1, c2 = st.columns(2)
                            with c1:
                                if sobras:
                                    st.write("### Vagas Remanescentes")
                                    st.dataframe(pd.DataFrame(list(sobras), columns=["Unidade"]), use_container_width=True)
                                else:
                                    st.info("Todas as vagas foram preenchidas.")
                            with c2:
                                st.write("### Exportação")
                                excel_data = gerar_excel_em_memoria(df_resultado, sobras)
                                st.download_button("📥 Baixar Planilha (.xlsx)", excel_data, "Resultado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        
                        with tab3:
                            st.text_area("Logs", value="\n".join(log), height=400)
                    else:
                        st.warning("Nenhuma movimentação gerada. Verifique se as vagas iniciais correspondem às opções.")