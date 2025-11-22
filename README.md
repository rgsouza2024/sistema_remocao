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

## 🎯 O Desafio

Nos concursos de remoção da magistratura, a movimentação de um juiz gera uma **vacância derivada** (a vaga que ele deixa para trás). Essa nova vaga deve ser ofertada imediatamente aos candidatos mais antigos que a desejavam, criando um efeito em cadeia complexo.

Realizar essa análise manualmente é propenso a erros humanos, pois exige reavaliar a lista de antiguidade inteira a cada nova vaga que surge no meio do processo ("Repescagem"), além de lidar com múltiplas opções de preferência por candidato.

## 💡 A Solução Tecnológica

Este software automatiza a leitura do edital e processa a alocação utilizando um algoritmo de **Restart on Move** (Reinício após Movimentação).

### Principais Funcionalidades:

* **🔄 Algoritmo de Repescagem:** Diferente de uma varredura linear simples, o sistema reinicia a verificação do topo da lista de antiguidade sempre que uma nova vaga é aberta. Isso garante matematicamente que o magistrado mais antigo tenha prioridade sobre qualquer vaga surgida.
* **📄 Ingestão de Dados (ETL):** Lê arquivos `.docx` (Word), extrai tabelas de inscritos, converte datas e normaliza nomes de unidades judiciárias.
* **📊 Visualização de Grafos:** Utiliza `NetworkX` e `PyVis` para gerar um arquivo HTML interativo (`Grafo_Remocao.html`), desenhando as conexões entre Origem e Destino.
* **📑 Relatórios em Excel:** Gera um arquivo `.xlsx` contendo a tabela oficial de resultados e a lista de vagas remanescentes.

## 🚀 Como Usar esta Aplicação

1.  Arraste o arquivo `.docx` com a tabela de inscritos para a área de upload.
2.  O sistema detectará automaticamente as vagas disponíveis no edital.
3.  Clique em **"Iniciar Processamento"**.
4.  Navegue pelas abas para ver o **Grafo Visual**, a **Tabela de Resultados** e baixar os relatórios.

---

## 👨‍💻 Autor

Desenvolvido por **Rodrigo Gonçalves de Souza**, juiz federal do TRF da 1ª Região.