# Automação de Exportação de Relatórios Power BI

Projeto de automação em Python para extração de dados do Power BI e integração com Google Sheets.

## 📌 Objetivo

Automatizar a exportação de dados de relatórios do Power BI, atualizar uma planilha no Google Sheets e permitir que um processo automatizado distribua oportunidades de vendas para vendedores.

Este projeto demonstra automação de processos de BI utilizando Python e integração entre ferramentas de análise de dados.

## ⚙️ Tecnologias utilizadas

* Python
* PyAutoGUI
* Pandas
* Google Sheets API
* gspread

## 🔄 Fluxo da automação

1. Abre automaticamente um relatório `.pbix`
2. Atualiza os dados do Power BI
3. Exporta dois relatórios em CSV
4. Remove arquivos antigos
5. Envia os dados para abas específicas de uma planilha Google Sheets
6. A planilha executa um script que distribui oportunidades de vendas automaticamente

## 📂 Estrutura do projeto

```
aut-relatorio-daily/
│
├─ src
│   └─ main.py
│
├─ config
│   └─ cred.json
│
├─ data
│
├─ tests
│
├─ config_example.py
├─ requirements.txt
├─ .gitignore
└─ README.md
```

## 🔧 Configuração

1. Clone o repositório

```
git clone <url-do-repositorio>
```

2. Instale as dependências

```
pip install -r requirements.txt
```

3. Copie o arquivo de configuração

```
config_example.py → config_local.py
```

4. Configure os caminhos no arquivo `config_local.py`

## 🔐 Segurança

Este repositório **não contém dados sensíveis**, incluin
