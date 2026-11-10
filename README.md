# 🍣 Restaurante Mazon — Gestão de Estoque & Inteligência Financeira (Fase 1)

Sistema web desenvolvido em **Python + Streamlit** com design system **Dark Luxury** (`#121212` / `#262626` / `#D4A24C`) e persistência em nuvem via **Google Sheets** e **Google Drive**.

---

## 📁 Estrutura do Projeto

- `app.py` — Aplicação principal Streamlit com os 5 módulos operacionais e gráficos interativos Plotly.
- `sheets_manager.py` — Módulo de conexão com Google Sheets (`gspread`), Google Drive API, criação automática das 5 abas obrigatórias e hooks de OCR/IA para a Fase 2.
- `style.css` — Folha de estilo Dark Luxury injetada no Streamlit.
- `requirements.txt` — Dependências Python do projeto.
- `credentials_example.json` — Modelo do arquivo de credenciais da Service Account do Google Cloud.

---

## ☁️ Passo a Passo: Criando a Service Account no Google Cloud

1. **Acesse o Google Cloud Console:**
   - Entre em https://console.cloud.google.com/ e crie um novo projeto (ex: `restaurante-mazon-erp`).

2. **Ative as APIs Obrigatórias:**
   - No menu lateral, vá em **APIs e Serviços > Biblioteca**.
   - Pesquise e clique em **Ativar** nas duas APIs abaixo:
     1. **Google Sheets API**
     2. **Google Drive API**

3. **Crie a Service Account (Conta de Serviço):**
   - Vá em **APIs e Serviços > Credenciais**.
   - Clique em **+ Criar Credenciais > Conta de serviço**.
   - Defina um nome (ex: `mazon-estoque-bot`) e clique em **Criar e Continuar**.
   - Em Papel (Role), selecione **Básico > Editor** e conclua.

4. **Gere e Baixe a Chave JSON (`credentials.json`):**
   - Na lista de Contas de Serviço, clique sobre o e-mail recém-criado (ex: `mazon-estoque-bot@restaurante-mazon-erp.iam.gserviceaccount.com`).
   - Acesse a aba **Chaves (Keys)** > **Adicionar Chave** > **Criar nova chave**.
   - Escolha o formato **JSON** e clique em **Criar**.
   - Renomeie o arquivo baixado para `credentials.json` e coloque-o na raiz desta pasta (`Estoque Mazon`).

5. **Compartilhe a Planilha no Google Sheets:**
   - Crie uma planilha no Google Sheets com o nome `Restaurante_Mazon_ERP_Database`.
   - Clique no botão **Compartilhar** (canto superior direito do Google Sheets) e adicione o `client_email` que está dentro do seu `credentials.json` com permissão de **Editor**.
   - *Nota:* Ao iniciar o sistema pela primeira vez conectado, o módulo `sheets_manager.py` cria e popula automaticamente as 5 abas obrigatórias (`Estoque`, `Fornecedores_Precos`, `Entradas_Compras`, `Saidas_Vendas`, `Ficha_Tecnica`) caso ainda não existam.

---

## 🚀 Como Executar o Projeto no Terminal

```powershell
# 1. Crie um ambiente virtual (opcional, mas recomendado)
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 2. Instale as dependências
pip install -r requirements.txt

# 3. Execute o servidor Streamlit
streamlit run app.py
```
