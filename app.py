import streamlit as st
import pandas as pd
from docx import Document
import networkx as nx
from pyvis.network import Network
from openpyxl.styles import Alignment
import io
import os
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
    /* Força a cor AZUL para botões do tipo 'primary' */
    div.stButton > button[kind="primary"] {
        background-color: #007BFF;
        color: white;
        border: none;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #0056b3; /* Azul mais escuro ao passar o mouse */
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

# --- FUNÇÕES DE EXTRAÇÃO E LÓGICA ---

def ler_arquivo_word(uploaded_file):
    """Lê o arquivo DOCX enviado pelo usuário."""
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
        st.error(f"Erro ao ler arquivo: {e}")
        return None

def detectar_vagas_do_edital(df):
    """
    Varre todas as opções dos candidatos procurando por 'DISPONÍVEL'.
    Retorna uma lista única de vagas ofertadas.
    """
    vagas_detectadas = set()
    colunas_opcoes = [col for col in df.columns if 'Opção' in col]
    
    for col in colunas_opcoes:
        valores = df[col].dropna().unique()
        for val in valores:
            if isinstance(val, str) and "DISPONÍVEL" in val.upper():
                nome_vaga = val.split(" - ")[0].strip()
                vagas_detectadas.add(nome_vaga)
                
    return list(vagas_detectadas)

def normalizar_nome(texto):
    if not isinstance(texto, str): return ""
    texto = texto.upper().strip()
    if " - " in texto:
        texto = texto.split(" - ")[0]
    return texto.strip()

def encurtar_nome(texto):
    if not texto: return "Origem Desconhecida"
    texto = normalizar_nome(texto)
    texto = texto.replace("DA SJ ", "").replace("DA SSJ ", "").replace("SUBSEÇÃO JUDICIÁRIA DE ", "")
    texto = texto.replace("SEÇÃO JUDICIÁRIA DE ", "").replace("SEÇÃO JUDICIÁRIA DO ", "")
    texto = texto.replace("VARA ÚNICA", "VARA UNICA").replace("RELATORIA DA ", "REL. ")
    if len(texto) > 35:
        return texto[:32] + "..."
    return texto

def processar_remocao(df, vagas_iniciais_lista):
    vagas_abertas = set([normalizar_nome(v) for v in vagas_iniciais_lista])
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
            lotacao_atual_norm = normalizar_nome(lotacao_atual)
            opcoes = juiz['Lista_Opcoes']
            
            match_encontrado = False
            vaga_escolhida = ""
            opcao_numero = 0
            opcao_final_texto = ""
            
            for i, opcao_bruta in enumerate(opcoes):
                opcao_norm = normalizar_nome(opcao_bruta)
                for vaga_aberta in vagas_abertas:
                    if vaga_aberta in opcao_norm:
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
        if not origem: origem = "Externo"
            
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
st.markdown("""
Esta aplicação processa a cadeia de vacância em concursos de remoção.
Faça upload do arquivo **.docx** e o sistema identificará automaticamente as vagas disponíveis.
""")

# 1. Upload
uploaded_file = st.file_uploader("Arraste o arquivo DOCX com os inscritos aqui", type="docx")

if uploaded_file is not None:
    df_bruto = ler_arquivo_word(uploaded_file)
    
    if df_bruto is not None:
        # --- AUTODETECÇÃO DE VAGAS ---
        vagas_detectadas = detectar_vagas_do_edital(df_bruto)
        texto_padrao = "\n".join(vagas_detectadas) if vagas_detectadas else ""
        
        st.info(f"🔎 O sistema detectou {len(vagas_detectadas)} vagas ofertadas no edital.")
        
        with st.expander("Ver ou Editar Vagas Iniciais Detectadas", expanded=True):
            texto_vagas_finais = st.text_area(
                "Vagas consideradas para o início da disputa (uma por linha):",
                value=texto_padrao,
                height=100
            )
        
        # Botão de Ação Principal (Agora AZUL via CSS)
        if st.button("🚀 Iniciar Processamento da Remoção", type="primary"):
            with st.spinner('Calculando antiguidade e cadeia de vacância...'):
                
                vagas_iniciais = [v.strip() for v in texto_vagas_finais.split('\n') if v.strip()]
                
                # Tratamento de Dados
                df_bruto['Data de Exercício'] = pd.to_datetime(df_bruto['Data de Exercício'], format='%d/%m/%Y', errors='coerce')
                df_ordenado = df_bruto.sort_values(by=['Data de Exercício', 'Matrícula'], ascending=[True, True])
                
                opcoes_limpas = []
                colunas_opcoes = [col for col in df_bruto.columns if 'Opção' in col]
                for _, row in df_ordenado.iterrows():
                    lista = [row[col] for col in colunas_opcoes if row[col] and str(row[col]).strip() != ""]
                    opcoes_limpas.append(lista)
                df_ordenado['Lista_Opcoes'] = opcoes_limpas
                
                # Motor de Decisão
                df_resultado, log, sobras = processar_remocao(df_ordenado, vagas_iniciais)
                
                if not df_resultado.empty:
                    st.success("✅ Análise concluída!")
                    
                    tab1, tab2, tab3 = st.tabs(["📊 Resultado Visual", "📋 Tabela Oficial", "📜 Logs Detalhados"])
                    
                    with tab1:
                        st.write("### Mapa da Cadeia de Vacância")
                        html_grafo = gerar_html_grafo(df_resultado)
                        components.html(html_grafo, height=650, scrolling=True)
                        st.download_button("📥 Baixar Grafo (HTML)", html_grafo, "Grafo_Remocao.html", "text/html")

                    with tab2:
                        st.write("### Resultado para Publicação")
                        st.dataframe(df_resultado[['Nome', 'Origem', 'Destino', 'Opção Nº']], use_container_width=True)
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.write("### Vagas Remanescentes")
                            if sobras:
                                st.dataframe(pd.DataFrame(list(sobras), columns=["Unidade"]), use_container_width=True)
                            else:
                                st.info("Todas as vagas foram preenchidas.")
                        
                        with col2:
                            st.write("### Exportação")
                            excel_data = gerar_excel_em_memoria(df_resultado, sobras)
                            st.download_button(
                                label="📥 Baixar Planilha Oficial (.xlsx)",
                                data=excel_data,
                                file_name="Resultado_Final_Repescagem.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                    with tab3:
                        st.text_area("Histórico de Movimentações", value="\n".join(log), height=400)
                
                else:
                    st.warning("Nenhuma movimentação ocorreu. Verifique se as vagas iniciais estão escritas exatamente como nas opções dos candidatos.")