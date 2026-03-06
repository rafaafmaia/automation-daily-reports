# 🤖 Automação de Relatórios Diários (Daily Reports)

Este projeto automatiza o fluxo completo de atualização e exportação de relatórios do **Power BI** para o **Google Sheets**. Ele utiliza RPA (Robotic Process Automation) para interagir com a interface do Power BI Desktop e integra APIs do Google para disponibilizar os dados na nuvem.

## 🛠️ Tecnologias Utilizadas

* **Python**: Linguagem base do projeto.
* **PyAutoGUI**: Automação da interface gráfica e controle de mouse/teclado.
* **Pandas**: Manipulação e tratamento de dados.
* **Gspread & Gspread-dataframe**: Integração robusta com a API do Google Sheets.
* **Google Cloud Service Account**: Gerenciamento seguro de credenciais e acessos.

## 📁 Estrutura do Projeto

* `src/`: Contém os scripts principais da automação.
* `config/`: Pasta para armazenar credenciais de API (protegida por `.gitignore`).
* `config_example.py`: Modelo de configuração para facilitar o deploy por outros desenvolvedores.
* `requirements.txt`: Lista de dependências para instalação rápida.

## 🔐 Segurança e Boas Práticas

Este repositório foi desenvolvido seguindo padrões de segurança cibernética:
* **Gestão de Segredos**: Arquivos de credenciais (`cred.json`) e configurações locais estão devidamente ignorados no sistema de versionamento.
* **Modularização de Configurações**: Uso de variáveis de ambiente e arquivos locais para evitar dados sensíveis no código-fonte.

## 🚀 Como Executar

1. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt

2. **Configure suas credenciais:**

Crie uma Service Account no Google Cloud.

Salve o JSON em config/cred.json.

Crie um arquivo src/config_local.py baseado no config_example.py com seus caminhos locais.

3. **Rode a automação:**
   Bash
   python src/main.py

Desenvolvido por Rafael Maia - https://www.linkedin.com/in/rafaafmaia/