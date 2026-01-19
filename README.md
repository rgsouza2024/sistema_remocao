---
title: Sistema de Analise de Remocao
emoji: ⚖️
colorFrom: blue
colorTo: green
sdk: streamlit
sdk_version: 1.40.1
app_file: app.py
pinned: false
license: mit
---

# ⚖️ Sistema de Análise de Remoção de Magistrados

> Algoritmo em Python para automação, análise e visualização de concursos de remoção de juízes federais, garantindo precisão rigorosa no critério de antiguidade, processamento da cadeia de vacância (efeito dominó) e aplicação de regras restritivas.

![Python](https://img.shields.io/badge/Python-3.12-blue)
![Data Viz](https://img.shields.io/badge/Visualização-PyVis%20%7C%20NetworkX-orange)
![Reporting](https://img.shields.io/badge/Relatórios-Excel%20%7C%20OpenPyXL-green)
![Version](https://img.shields.io/badge/Versão-13.0-brightgreen)

---

## 📸 Visualização da Cadeia de Vacância

O sistema gera um **grafo interativo** que permite visualizar a "Dança das Cadeiras": quem ocupou a vaga de quem e qual unidade foi liberada na sequência.

![Grafo de Remoção](preview_grafo.png)

---

## 🎯 O Desafio

Nos concursos de remoção da magistratura, a movimentação de um juiz gera uma **vacância derivada** (a vaga que ele deixa para trás). O processo envolve complexidades logísticas como:

1.  **Efeito Dominó:** A vaga deixada por um juiz deve ser ofertada imediatamente aos candidatos mais antigos.
2.  **Candidato Sôfrego (Upgrade):** Se uma vaga prioritária surge tardiamente, o magistrado mais antigo deve ter o direito de trocar sua escolha anterior por esta melhor.
3.  **Regras de Bloqueio (Congelamento):** Magistrados removidos recentemente (ex: há menos de 1 ano) podem sofrer penalidades na ordem de classificação.

---

## 🧠 Arquitetura do Motor de Remoção (V13)

O núcleo do sistema é o algoritmo `processar_remocao()`, que implementa as regras oficiais do TRF para concursos de remoção.

### Regras de Negócio Implementadas

| Regra | Descrição | Implementação |
|-------|-----------|---------------|
| **Antiguidade Soberana** | O candidato mais antigo (por Data de Exercício + Matrícula) sempre tem prioridade de escolha | Ordenação do DataFrame antes do processamento |
| **Restart on Move** | A cada movimentação, o algoritmo reinicia do candidato mais antigo | Loop `while houve_movimentacao` com `break` |
| **Lookahead Anti-Bloqueio** | O sênior pode ceder a vez se pegar uma vaga bloquearia sua 1ª opção | Função `detectar_bloqueio()` |
| **Smart Match** | Comparação de nomes de vagas usando word boundaries (regex) | Função `smart_match()` |

### Fluxograma do Algoritmo

```
┌─────────────────────────────────────────────────────────────────┐
│                    INÍCIO DO PROCESSAMENTO                       │
├─────────────────────────────────────────────────────────────────┤
│  1. Ordenar candidatos por [Data de Exercício, Matrícula]       │
│  2. Inicializar vagas_abertas com vagas do edital               │
│  3. ciclo = 0                                                    │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    LOOP PRINCIPAL (while)                        │
├─────────────────────────────────────────────────────────────────┤
│  houve_movimentacao = False                                      │
│  ciclo += 1                                                      │
│                                                                  │
│  PARA cada candidato em ordem de antiguidade:                   │
│    │                                                             │
│    ├─► Buscar primeira opção disponível em vagas_abertas        │
│    │                                                             │
│    ├─► SE encontrou match:                                       │
│    │     │                                                       │
│    │     ├─► LOOKAHEAD: detectar_bloqueio()?                    │
│    │     │     │                                                 │
│    │     │     ├─► SIM: Ceder vez (log "⏸️ cedeu vez")         │
│    │     │     │                                                 │
│    │     │     └─► NÃO: Confirmar remoção                       │
│    │     │           ├─► Adicionar lotação atual às vagas       │
│    │     │           ├─► Remover vaga escolhida                 │
│    │     │           ├─► houve_movimentacao = True              │
│    │     │           └─► BREAK (restart loop)                   │
│    │     │                                                       │
│    └─────┴───────────────────────────────────────────────────────│
│                                                                  │
│  SE houve_movimentacao == False: SAIR DO LOOP                   │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    FIM - Retornar resultados                     │
└─────────────────────────────────────────────────────────────────┘
```

### Funções Auxiliares do Motor

#### `smart_match(vaga, opcao) → bool`

Compara strings usando **word boundaries** (regex `\b`) para evitar matches incorretos.

```python
# Exemplo de problema resolvido:
"7A VARA" in "27A VARA"  # True ❌ (substring match incorreto)
smart_match("7A VARA", "27A VARA")  # False ✅ (word boundary correto)
```

**Implementação:**
```python
def smart_match(vaga, opcao):
    pattern = r'\b' + re.escape(vaga) + r'\b'
    return bool(re.search(pattern, opcao))
```

#### `detectar_bloqueio(senior, vaga_pretendida, indice) → bool`

Implementa o **Lookahead Anti-Bloqueio** verificando se o sênior bloquearia sua própria 1ª opção ao pegar uma vaga de menor preferência.

**Lógica:**
1. Se `indice == 0` (é a 1ª opção), não há bloqueio possível
2. Para cada opção MELHOR (índice menor):
   - Identificar quem ocupa essa vaga (`get_ocupante_lotacao()`)
   - Verificar se o ocupante quer a `vaga_pretendida`
   - Se SIM → **BLOQUEIO DETECTADO** (sênior deve ceder)

**Exemplo Prático:**
```
BRUNO (sênior): 1ª opção = 27ª Vara, 3ª opção = 7ª Vara
JOSÉ MÁRCIO (júnior): Ocupa 27ª Vara, quer 7ª Vara como 2ª opção

Quando 7ª Vara abre:
├─► BRUNO vê 7ª Vara disponível (sua 3ª opção)
├─► Lookahead: "Quem ocupa minha 1ª opção (27ª Vara)?"
│   └─► JOSÉ MÁRCIO ocupa
├─► Lookahead: "JOSÉ MÁRCIO quer a 7ª Vara?"
│   └─► SIM
├─► BLOQUEIO DETECTADO!
└─► BRUNO cede vez → JOSÉ MÁRCIO pega 7ª Vara → Abre 27ª Vara
    └─► BRUNO pega 27ª Vara (sua 1ª opção!) ✅
```

#### `padronizar_texto(texto) → str`

Normaliza strings para comparação uniforme, tratando:

| Transformação | Exemplo |
|---------------|---------|
| Maiúsculas | `"7ª Vara"` → `"7ª VARA"` |
| Símbolos ordinais | `"7ª"` → `"7A"` |
| Acentos (NFD) | `"Muriaé"` → `"MURIAE"` |
| Prefixos jurisdicionais | `"DA SJ"`, `"DA SSJ"` → removidos |
| Espaços múltiplos | `"VARA  ÚNICA"` → `"VARA UNICA"` |
| Sufixos | `"- DISPONÍVEL"` → removido |

---

## 💡 Principais Funcionalidades

### Motor de Remoção
* **🔍 Lookahead Anti-Bloqueio (V13):** Evita que o candidato mais antigo "dê um tiro no próprio pé" ao pegar uma vaga que bloquearia sua opção preferencial.
* **❄️ Regra de Congelamento:** Interface para seleção manual de magistrados penalizados por remoção recente.
* **🔄 Lógica de Upgrade:** Garantia de que o magistrado sempre obtenha a melhor vaga possível.
* **🛡️ Smart Match:** Comparação precisa de nomes de varas usando regex word boundaries.

### Interface e Usabilidade
* **🏛️ Identidade Institucional:** Logo do TRF1 no cabeçalho com cores institucionais (#002F6C).
* **📋 Abas Reorganizadas:** "Quadro de Remoções" como aba padrão, facilitando acesso ao resultado principal.
* **✅ Validação de Colunas:** Verificação prévia de colunas obrigatórias com mensagens de erro amigáveis.
* **✅ Validação de Colunas:** Verificação prévia de colunas obrigatórias com mensagens de erro amigáveis.
* **📂 Suporte Multi-Formato:** Leitura de arquivos **Word (.docx)**, **Excel (.xlsx)** e **JSON (.json)**.
* **🔤 Normalização Robusta:** Reconhecimento inteligente de colunas independente de acentos ou maiúsculas (ex: "Início da Lotação" = "inicio da lotacao").
* **❄️ Congelamento Inteligente:** Detecção automática baseada na palavra-chave "remoção" (substring case-insensitive).
* **📊 Visualização de Grafos:** Mapa visual interativo das movimentações com PyVis (física otimizada).
* **📑 Relatórios Oficiais:** Planilha `.xlsx` formatada com resultado final e vagas remanescentes.

---

## 🚀 Como Usar

1. Acesse a interface web Streamlit.
2. Arraste o arquivo de inscritos (**DOCX** ou **XLSX**) para a área de upload.
3. Confira as vagas iniciais detectadas automaticamente.
4. **(Opcional)** Marque **"Existem inscritos congelados?"** e selecione os nomes.
5. Clique em **"Iniciar Processamento"**.
6. Navegue pelas abas: **Quadro de Remoções**, **Resultado Visual**, **Logs**.

---

## 🛠️ Stack Tecnológico

| Tecnologia | Uso |
|------------|-----|
| **Python 3.12** | Linguagem base |
| **Streamlit** | Interface Web interativa |
| **Pandas** | Manipulação de dados e ordenação |
| **Python-Docx** | Leitura de arquivos Word |
| **OpenPyXL** | Leitura/escrita de arquivos Excel |
| **NetworkX + PyVis** | Renderização de grafos interativos |
| **Regex (re)** | Smart Match com word boundaries |

---

## 📁 Estrutura do Projeto

```
sistema_remocao/
├── app.py              # Aplicação principal Streamlit (Motor V13)
├── analise_remocao.py  # Script standalone (versão CLI)
├── logo_trf1.png       # Logo institucional do TRF1
├── requirements.txt    # Dependências Python
├── README.md           # Esta documentação
└── lib/                # Bibliotecas auxiliares (vis.js, tom-select)
```

---

## 🔧 Instalação Local

```bash
# Clone o repositório
git clone https://github.com/rgsouza2024/sistema_remocao.git
cd sistema_remocao

# Instale as dependências
pip install -r requirements.txt

# Execute a aplicação
streamlit run app.py
```

---

## 📝 Changelog

### V13.2 (2026-01-19) - Robustez e Novos Formatos
- ✨ **Suporte a JSON:** Agora aceita arquivos `.json` (lista de objetos) além de Excel e Word.
- ✨ **Normalização de Colunas:** O sistema agora é "Case & Accent Insensitive". Aceita "Início da Lotação", "inicio da lotacao", "INICIO_LOTA", etc.
- ✨ **Congelamento Automático:** Regra de congelamento baseada na palavra-chave "remocao" (substring). Detecta "Remoção a Pedido", "REMOÇÃO", etc.
- ✨ **Detecção de Opções:** Normalização automática de colunas de opções (ex: "1ª OPÇÃO" -> "1ª Opção").

### V13.1 (2026-01-16) - Melhorias de UI/UX
- 🎨 **Identidade Institucional:** Adicionado logo do TRF1 no cabeçalho
- 🎨 **Cor Institucional:** Título em azul oficial (#002F6C)
- 📋 **Abas Reorganizadas:** "Quadro de Remoções" como aba principal (antes era "Resultado Visual")
- ⚙️ **Expander Colapsado:** Configurações do Edital iniciam recolhidas por padrão
- ✅ **Validação de Colunas:** Erro amigável quando colunas obrigatórias estão faltando
- 📊 **Grafo Otimizado:** Ajustes na física (spring_length=300) para melhor visualização

### V13 (2026-01-16)
- ✨ **Lookahead Anti-Bloqueio:** Novo algoritmo que detecta deadlocks e permite que o sênior ceda a vez estrategicamente
- ✨ **Smart Match:** Comparação com word boundaries evita match incorreto, por exemplo, entre "7A VARA" e "27A VARA"
- 🔧 Funções auxiliares: `detectar_bloqueio()`, `get_ocupante_lotacao()`, `smart_match()`

### V12 (2026-01-16)
- 🔧 Regra de Congelamento com seleção manual de magistrados

### V11
- 🔧 Lógica de Upgrade (troca de vaga quando surge opção melhor)
- 🔧 Algoritmo Restart on Move

---

## 👨‍💻 Autor

Desenvolvido por **Rodrigo Gonçalves de Souza** - Juiz Federal do TRF da 1ª Região.

---

## 📄 Licença

MIT License - Veja o arquivo LICENSE para detalhes.