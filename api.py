# api.py
# Interface API via FastAPI
# Expondo o motor de remoção para integração com outros sistemas

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse
from typing import Optional, List
import pandas as pd
import io
import json
from datetime import datetime

# Importando o motor "Cérebro"
from motor_remocao import (
    ler_arquivo_word, 
    ler_arquivo_excel, 
    ler_arquivo_json,
    normalizar_colunas, 
    detectar_vagas_do_edital, 
    remove_acentos,
    processar_remocao,
    aplicar_congelamento
)

app = FastAPI(
    title="API Sistema de Remoção",
    description="API para processamento de concursos de remoção de magistrados.",
    version="1.0.0"
)

@app.get("/")
def read_root():
    return {"status": "online", "message": "API de Remoção Operacional"}

@app.post("/analisar-remocao")
async def analisar_remocao(
    file: UploadFile = File(...),
    data_referencia: Optional[str] = Form(None, description="Data de referência no formato YYYY-MM-DD ou DD/MM/YYYY"),
    vagas_edital: Optional[str] = Form(None, description="Lista de vagas separadas por ponto e vírgula ou JSON array")
):
    """
    Recebe um arquivo (DOCX, XLSX, JSON) e parâmetros opcionais.
    Processa a remoção e retorna o resultado em JSON.
    """
    
    # 1. Leitura do Arquivo
    contents = await file.read()
    file_obj = io.BytesIO(contents)
    df_bruto = None
    
    filename = file.filename.lower()
    
    if filename.endswith('.docx'):
        df_bruto = ler_arquivo_word(file_obj)
    elif filename.endswith('.xlsx'):
        df_bruto = ler_arquivo_excel(file_obj)
    elif filename.endswith('.json'):
        # JSON file needs text decoding or explicit loading depending on implementation
        # ler_arquivo_json expects a file-like object containing text or bytes
        # Since we read bytes, io.BytesIO is fine for json.load
        df_bruto = ler_arquivo_json(file_obj)
    else:
        raise HTTPException(status_code=400, detail="Formato de arquivo não suportado. Use .docx, .xlsx ou .json")
    
    if df_bruto is None or df_bruto.empty:
        raise HTTPException(status_code=400, detail="Não foi possível ler dados do arquivo enviado.")

    # 2. Normalização
    df_bruto = normalizar_colunas(df_bruto)
    
    # Validação Básica
    colunas_obrigatorias = ['Nome', 'Matrícula', 'Lotação Atual', 'Data de Exercício']
    colunas_faltando = [col for col in colunas_obrigatorias if col not in df_bruto.columns]
    if colunas_faltando:
        raise HTTPException(status_code=400, detail=f"Colunas obrigatórias faltando: {', '.join(colunas_faltando)}")

    # 3. Tratamento de Vagas
    lista_vagas = []
    if vagas_edital:
        # Tenta interpretar como JSON array ou string separada por ; ou \n
        try:
            lista_vagas = json.loads(vagas_edital)
            if not isinstance(lista_vagas, list):
                lista_vagas = str(vagas_edital).replace(';', '\n').split('\n')
        except:
            lista_vagas = str(vagas_edital).replace(';', '\n').split('\n')
    else:
        # Detecta automático se não informado
        lista_vagas = detecting_vacancies = detecting_vacancies = detectar_vagas_do_edital(df_bruto)
    
    # Limpeza lista vagas
    lista_vagas = [str(v).strip() for v in lista_vagas if str(v).strip()]

    # 4. Tratamento de Datas
    df_bruto['Data de Exercício'] = pd.to_datetime(df_bruto['Data de Exercício'], dayfirst=True, errors='coerce')
    if df_bruto['Data de Exercício'].isna().all():
         raise HTTPException(status_code=400, detail="Erro: Datas de exercício inválidas ou formato não reconhecido.")

    # 5. Congelamento Automático
    # Define data de referência (default = hoje)
    dt_ref = pd.Timestamp.now()
    if data_referencia:
        try:
            # Tenta formatos comuns
            dt_ref = pd.to_datetime(data_referencia, dayfirst=True) 
        except:
            pass # Mantém hoje se falhar
            
    df_bruto, congelados, incompletos = aplicar_congelamento(df_bruto, dt_ref)

    # 6. Ordenação
    # Status_Congelado (0 primeiro, 1 depois) + Antiguidade
    criterios_ordenacao = ['Status_Congelado', 'Data de Exercício', 'Matrícula']
    ascendencia = [True, True, True]
    df_ordenado = df_bruto.sort_values(by=criterios_ordenacao, ascending=ascendencia)

    # 7. Prepara Lista de Opções
    opcoes_limpas = []
    colunas_opcoes = [col for col in df_bruto.columns if 'Opção' in col]
    for _, row in df_ordenado.iterrows():
        lista = [row[col] for col in colunas_opcoes if row[col] and str(row[col]).strip() not in ["", "nan"]]
        opcoes_limpas.append(lista)
    df_ordenado['Lista_Opcoes'] = opcoes_limpas

    # 8. PROCESSAMENTO CORE
    df_resultado, log, sobras = processar_remocao(df_ordenado, lista_vagas)

    # 9. Retorno JSON
    resultado_json = df_resultado.to_dict(orient='records')
    
    response_data = {
        "status": "sucesso",
        "data_referencia_utilizada": dt_ref.strftime("%Y-%m-%d"),
        "total_movimentacoes": len(df_resultado),
        "vagas_remanescentes": list(sobras),
        "magistrados_congelados": congelados,
        "alertas_integridade": incompletos,
        "movimentacoes": resultado_json,
        "logs": log
    }
    
    return JSONResponse(content=response_data)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
