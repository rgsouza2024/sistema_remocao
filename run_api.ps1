
Write-Host "Iniciando API de Remoção..." -ForegroundColor Cyan
Write-Host "Acesse a documentação em: http://127.0.0.1:8000/docs" -ForegroundColor Green

uvicorn api:app --reload --host 127.0.0.1 --port 8000
