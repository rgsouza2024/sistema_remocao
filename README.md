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

## 💡 A Solução Tecnológica

Este software automatiza a leitura do edital e processa a alocação utilizando um algoritmo de **Restart on Move** (Reinício após Movimentação), com camadas de regras de negócio configuráveis.

### Principais Funcionalidades:

* **❄️ Regra de Congelamento (Novo):** Interface para seleção manual de magistrados que devem ser "congelados" (penalizados por remoção recente). O algoritmo reordena a lista automaticamente, movendo-os para o final da fila de prioridade.
* **🔄 Lógica de Upgrade:** O sistema garante que o magistrado sempre obtenha a melhor vaga possível dentro de suas opções. Se uma vaga de maior preferência surgir após ele já ter sido alocado, o sistema realiza a troca e devolve a vaga anterior ao pote.
* **📂 Suporte Híbrido:** Leitura transparente de arquivos **Word (.docx)** e **Excel (.xlsx)** com tratamento de dados unificado.
* **🛡️ Sanitização de Dados:** Normalização avançada que corrige divergências de digitação (ex: "22ª Vara" vs "22A VARA"), acentuação e espaços invisíveis.
* **📅 Tratamento de Datas (PT-BR):** Conversão segura de datas de exercício, prevenindo erros de inversão dia/mês.
* **📊 Visualização de Grafos:** Gera um mapa visual interativo das movimentações.
* **📑 Relatórios Oficiais:** Gera planilha `.xlsx` formatada com o resultado final e vagas remanescentes.

## 🚀 Como Usar esta Aplicação

1.  Acesse a interface web.
2.  Arraste o arquivo de inscritos (**DOCX** ou **XLSX**) para a área de upload.
3.  Confira as vagas iniciais detectadas automaticamente.
4.  **(Opcional)** Marque a opção **"Existem inscritos congelados?"** e selecione os nomes na lista.
5.  Clique em **"Iniciar Processamento"**.
6.  Navegue pelas abas para ver o **Grafo Visual**, a **Tabela de Resultados** e baixar os relatórios.

## 🛠️ Stack Tecnológico

* **Python 3.12**
* **Streamlit:** Interface Web interativa.
* **Pandas:** Manipulação de dados e ordenação temporal.
* **Python-Docx & OpenPyXL:** Leitura e escrita de arquivos Office.
* **NetworkX + PyVis:** Renderização gráfica de redes.

---

## 👨‍💻 Autor

Desenvolvido por **Rodrigo Gonçalves de Souza** - juiz federal do TRF da 1ª Região.