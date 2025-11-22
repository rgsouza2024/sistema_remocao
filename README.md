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

> Algoritmo em Python para automação, análise e visualização de concursos de remoção de juízes federais, garantindo precisão rigorosa no critério de antiguidade e processamento da cadeia de vacância (efeito dominó).

![Python](https://img.shields.io/badge/Python-3.12-blue)
![Data Viz](https://img.shields.io/badge/Visualização-PyVis%20%7C%20NetworkX-orange)
![Reporting](https://img.shields.io/badge/Relatórios-Excel%20%7C%20OpenPyXL-green)

---

## 📸 Visualização da Cadeia de Vacância

O sistema gera um **grafo interativo** que permite visualizar a "Dança das Cadeiras": quem ocupou a vaga de quem e qual unidade foi liberada na sequência.

![Grafo de Remoção](preview_grafo.png)

---

## 🎯 O Desafio

Nos concursos de remoção da magistratura, a movimentação de um juiz gera uma **vacância derivada** (a vaga que ele deixa para trás). Essa nova vaga deve ser ofertada imediatamente aos candidatos mais antigos que a desejavam, criando um efeito em cadeia complexo.

Realizar essa análise manualmente é propenso a erros humanos devido a:
1.  **Volume de Dados:** Múltiplas opções de preferência por candidato.
2.  **Dinâmica de Repescagem:** A necessidade de reavaliar a lista de antiguidade inteira a cada nova vaga que surge.
3.  **Inconsistência de Formatos:** Dados vindos de fontes diferentes (Word vs Excel) com formatações de texto divergentes.

## 💡 A Solução Tecnológica

Este software automatiza a leitura do edital e processa a alocação utilizando um algoritmo de **Restart on Move** (Reinício após Movimentação), com uma camada robusta de tratamento de dados.

### Principais Funcionalidades:

* **🔄 Algoritmo de Repescagem:** O sistema reinicia a verificação do topo da lista de antiguidade sempre que uma nova vaga é aberta, garantindo matematicamente a preferência do magistrado mais antigo.
* **📂 Suporte Multi-formato:** Leitura transparente de arquivos **Word (.docx)** e **Excel (.xlsx)**.
* **🛡️ Sanitização de Dados (Data Cleaning):** Normalização avançada de texto que remove acentos, padroniza ordinais (ª/º), corrige traços e espaços invisíveis, garantindo que "22ª Vara" (Excel) e "22A VARA" (Word) sejam interpretadas idênticamente.
* **📅 Tratamento de Datas (PT-BR):** Conversão segura de datas de exercício, prevenindo erros de inversão dia/mês comuns em planilhas.
* **📊 Visualização de Grafos:** Gera um mapa visual interativo (`Grafo_Remocao.html`) das movimentações.
* **📑 Relatórios Oficiais:** Gera planilha `.xlsx` formatada com o resultado final e vagas remanescentes.

## 🚀 Como Usar esta Aplicação

1.  Acesse a interface web.
2.  Arraste o arquivo de inscritos (**DOCX** ou **XLSX**) para a área de upload.
3.  O sistema detectará automaticamente as vagas disponíveis no edital.
4.  Clique em **"Iniciar Processamento"**.
5.  Navegue pelas abas para ver o **Grafo Visual**, a **Tabela de Resultados** e baixar os relatórios.

## 🛠️ Stack Tecnológico

* **Python 3.12**
* **Streamlit:** Interface Web interativa.
* **Pandas:** Manipulação de dados e ordenação temporal.
* **Python-Docx & OpenPyXL:** Leitura e escrita de arquivos Office.
* **NetworkX + PyVis:** Renderização gráfica de redes.

---

## 👨‍💻 Autor

Desenvolvido por **Rodrigo Gonçalves de Souza**, juiz federal do TRF da 1ª Região.