
from fastapi.testclient import TestClient
from api import app
import json
import io

client = TestClient(app)

def test_api_json_upload():
    # 1. Mock Data (JSON)
    mock_data = [
        {
            "Nome": "Juiz Teste 1",
            "Matrícula": "1001",
            "Lotação Atual": "1ª Vara de Teste",
            "Data de Exercício": "01/01/2020",
            "início da lotação": "01/01/2025",
            "motivo da lotação": "Remoção",
            "1ª Opção": "2ª Vara de Teste",
            "2ª Opção": "3ª Vara de Teste"
        },
        {
            "Nome": "Juiz Teste 2",
            "Matrícula": "1002",
            "Lotação Atual": "2ª Vara de Teste",
            "Data de Exercício": "01/01/2019", # Mais antigo
            "início da lotação": "01/01/2020",
            "motivo da lotação": "Promoção",
            "1ª Opção": "3ª Vara de Teste"
        }
    ]
    
    # 2. Convert to Bytes for Upload
    json_content = json.dumps(mock_data).encode('utf-8')
    files = {'file': ('dados.json', json_content, 'application/json')}
    
    # 3. Parameters
    data_ref = "30/01/2026"
    vagas = "2ª Vara de Teste;3ª Vara de Teste;4ª Vara de Teste"
    
    # 4. Request
    print("Enviando requisição POST para /analisar-remocao...")
    response = client.post(
        "/analisar-remocao",
        files=files,
        data={
            "data_referencia": data_ref,
            "vagas_edital": vagas
        }
    )
    
    # 5. Validation
    print(f"Status Code: {response.status_code}")
    if response.status_code == 200:
        data = response.json()
        print("✅ Sucesso! Resposta JSON recebida.")
        print(f"Status da API: {data.get('status')}")
        print(f"Total Movimentações: {data.get('total_movimentacoes')}")
        print(f"Congelados: {data.get('magistrados_congelados')}")
        
        # Verify freezing logic (Juiz Teste 1 should be frozen because date is recent + removal)
        # 01/01/2025 to 30/01/2026 is actually > 1 year (365 days).
        # Wait, 2025 to 2026 is > 1 year. Let's make it frozen.
        # Set reference date to 02/01/2025 so diff is small.
        
    else:
        print("❌ Falha na requisição.")
        print(response.text)

def test_freezing_logic_api():
    # Make Juiz 1 frozen (Removal < 1 year)
    mock_data = [
        {
            "Nome": "Juiz Congelado",
            "Matrícula": "999",
            "Lotação Atual": "Origem",
            "Data de Exercício": "01/01/2010",
            "início da lotação": "01/01/2026", # Very recent
            "motivo da lotação": "REMOÇÃO",
            "1ª Opção": "Destino"
        }
    ]
    json_content = json.dumps(mock_data).encode('utf-8')
    files = {'file': ('dados.json', json_content, 'application/json')}
    
    response = client.post(
        "/analisar-remocao",
        files=files,
        data={"data_referencia": "30/01/2026", "vagas_edital": "Destino"}
    )
    
    data = response.json()
    congelados = data.get('magistrados_congelados', [])
    print(f"Teste Congelamento: {congelados}")
    if any("Juiz Congelado" in s for s in congelados):
        print("✅ Lógica de congelamento funcionou na API!")
    else:
        print("❌ Lógica de congelamento FALHOU na API.")

if __name__ == "__main__":
    test_api_json_upload()
    test_freezing_logic_api()
