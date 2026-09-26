---
title: Sistema de Análise de Remoção
emoji: ⚖️
colorFrom: blue
colorTo: green
sdk: streamlit
sdk_version: 1.40.1
app_file: app.py
pinned: false
license: mit
---

# Sistema de Análise de Remoção de Magistrados

Aplicação para simular a distribuição de vagas em concursos de remoção de magistrados. Lê relações de inscritos, aplica critérios de antiguidade e preferências, acompanha as vagas abertas durante as movimentações e apresenta o resultado para conferência.

## Acesso e repositório

- **Aplicação publicada:** [Sistema de Remoção no Hugging Face Spaces](https://huggingface.co/spaces/rgsouza2024/sistema-remocao)
- **Código-fonte:** [repositório no GitHub](https://github.com/rgsouza2024/sistema_remocao), branch `main`
- **Interface publicada:** Streamlit, com `app.py` como arquivo de entrada do Space.

Em 25/09/2026, a branch `main` do Space e a branch `main` do GitHub apontavam para o mesmo commit (`33cd0c6`). Essa verificação confirma que estavam sincronizadas naquela data; não confirma sincronização automática para alterações futuras.

## Funcionalidades

- Importa arquivos Word (`.docx`), Excel (`.xlsx`) e JSON (`.json`).
- Detecta nomes de colunas comuns mesmo com diferenças de maiúsculas e acentos.
- Detecta vagas indicadas como disponíveis nas opções ou permite informar a lista manualmente.
- Aplica a regra automática de congelamento com base no motivo e na data de início da lotação.
- Processa movimentações em ciclos, reabrindo a lotação de origem e permitindo que candidatos melhorem uma alocação anterior quando uma opção melhor fica disponível.
- Exibe o quadro de remoções, vagas remanescentes e logs; a interface também permite baixar uma planilha Excel.
- Mostra as cadeias de remoção: cada cadeia parte de uma vaga do edital e segue pelas lotações abertas por quem se removeu, até a vaga remanescente. Movimentações que não partem do edital aparecem como permutas. A montagem pressupõe um único magistrado por lotação, o que vale porque titulares e substitutos concorrem em concursos separados.

## Regras e dados de entrada

O arquivo precisa conter estas colunas:

| Campo | Uso |
| --- | --- |
| `Nome` | Identificação do magistrado |
| `Matrícula` | Desempate na ordenação por antiguidade |
| `Lotação Atual` | Origem da movimentação e vaga que poderá ser aberta |
| `Data de Exercício` | Critério principal de antiguidade |

As colunas de opção devem indicar a posição, como `1ª Opção`, `2ª Opção` e assim por diante. O normalizador reconhece algumas variantes de nomes de colunas.

Na interface, as vagas iniciais são detectadas procurando `DISPONÍVEL` nas colunas de opção. A lista pode ser editada antes do processamento. No endpoint da API, a lista também pode ser enviada manualmente.

### Congelamento automático

O sistema marca como congelado o magistrado cujo motivo da lotação contém “remoção” e cuja lotação começou há menos de 365 dias em relação à data de referência. Dados de lotação ausentes ou inválidos não geram congelamento; quando identificados pelo motor, são incluídos nos alertas de integridade.

No comportamento atual, o congelamento altera a prioridade: os candidatos congelados ficam depois dos demais na ordenação, mas ainda podem ser processados se houver vaga compatível. A regra não os exclui automaticamente do concurso.

O motor procura a primeira opção disponível, verifica possíveis bloqueios e reinicia a análise após cada movimentação. A comparação dos nomes das vagas normaliza acentos e usa limites de palavra para evitar correspondências parciais como `7ª Vara` dentro de `27ª Vara`.

## Executar a interface localmente

Requisitos: Python 3.12 e acesso à internet para instalar as dependências.

A versão do Streamlit está fixada em `1.40.1` no `requirements.txt`, igual à versão declarada em `sdk_version` no cabeçalho deste README. Isso alinha as configurações do projeto; a versão efetivamente instalada no Space depende do commit implantado e deve ser conferida nos logs de build.

No PowerShell:

```powershell
git clone https://github.com/rgsouza2024/sistema_remocao.git
cd sistema_remocao
py -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m streamlit --version
.venv\Scripts\python.exe -m streamlit run app.py
```

O comando de versão deve informar `Streamlit, version 1.40.1`. O Streamlit informa no terminal o endereço local da aplicação. Use o Python do `.venv` para iniciar a interface com as dependências fixadas pelo projeto.

## API REST

A API FastAPI compartilha o motor de processamento com a interface.

### Executar com Docker

Na raiz do repositório:

```bash
docker build -t sistema-remocao-api .
docker run --rm -p 8000:8000 sistema-remocao-api
```

Depois, abra [`http://localhost:8000/docs`](http://localhost:8000/docs) para consultar e experimentar a documentação Swagger.

Para execução local no Windows, instale as dependências da API e rode o script de desenvolvimento:

```powershell
python -m pip install -r requirements_api.txt
.\run_api.ps1
```

O script usa `uvicorn --reload` e escuta apenas em `127.0.0.1`; é destinado ao desenvolvimento local.

### Endpoints

| Método e caminho | Descrição |
| --- | --- |
| `GET /` | Retorna um status simples da API |
| `POST /analisar-remocao` | Recebe um arquivo e retorna a análise em JSON |

O `POST /analisar-remocao` recebe `multipart/form-data`:

| Campo | Obrigatório | Descrição |
| --- | --- | --- |
| `file` | Sim | Arquivo `.docx`, `.xlsx` ou `.json` |
| `data_referencia` | Não | Data no formato `YYYY-MM-DD` ou `DD/MM/YYYY`; se omitida, usa a data atual |
| `vagas_edital` | Não | Vagas separadas por ponto e vírgula ou array JSON; se omitido, tenta detectar as vagas no arquivo |

Exemplo com `curl.exe` no PowerShell:

```powershell
curl.exe -X POST "http://127.0.0.1:8000/analisar-remocao" `
  -F "file=@inscritos.xlsx" `
  -F "data_referencia=30/01/2026" `
  -F "vagas_edital=7ª Vara;8ª Vara"
```

A resposta inclui `total_movimentacoes`, `movimentacoes`, `vagas_remanescentes`, `magistrados_congelados`, `alertas_integridade`, `logs` e `data_referencia_utilizada`. Cada item de `movimentacoes` traz também `Vaga`, a chave normalizada da vaga ocupada. A API retorna os dados em JSON; o download do Excel e as cadeias de remoção estão disponíveis na interface Streamlit.

### Observações de implantação

- A API não implementa autenticação nem limite próprio de tamanho para arquivos. Antes de disponibilizá-la fora de um ambiente local ou controlado, configure controle de acesso e limites na infraestrutura que a publica.
- No código atual, se `data_referencia` for informada em formato inválido, a API mantém a data atual como padrão. Confira `data_referencia_utilizada` na resposta.

## Estrutura principal

| Arquivo | Responsabilidade |
| --- | --- |
| `app.py` | Interface Streamlit e fluxo de interação |
| `api.py` | Endpoints FastAPI e validações da requisição |
| `motor_remocao.py` | Leitura, normalização, congelamento, processamento e geração de resultados |
| `requirements.txt` | Dependências da interface Streamlit e do motor |
| `requirements_api.txt` | Dependências da API e do motor |
| `Dockerfile` | Imagem e inicialização da API |

O script standalone legado `analise_remocao.py` foi removido. Para executar o sistema, use `app.py` (Streamlit) ou `api.py` (FastAPI); ambos compartilham o motor em `motor_remocao.py`.

## Testes

Os testes de `tests/` usam `unittest` e cobrem a montagem das cadeias de remoção:

```powershell
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Os antigos `test_api.py` e `test_refactor.py` foram removidos porque apenas imprimiam resultados e não validavam as respostas com asserções. A API pode ser verificada manualmente pela documentação em `/docs`.

## Licença

Os metadados deste Space declaram licença MIT. Este repositório não contém atualmente um arquivo `LICENSE` com o texto integral da licença.
