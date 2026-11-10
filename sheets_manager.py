"""
Módulo de Persistência em Nuvem (Google Sheets & Google Drive) - Restaurante Mazon
Gerencia autenticação via Service Account / OAuth, estrutura das 5 abas obrigatórias,
sincronização da base oficial de insumos, upload no Google Drive e preparação para OCR/IA.
"""

from __future__ import annotations

import io
import os
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# Escopos exigidos para Google Sheets e Google Drive
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Esquema obrigatório das 5 abas do banco de dados no Google Sheets
SHEETS_SCHEMA: Dict[str, List[str]] = {
    "Estoque": [
        "ID",
        "Item",
        "Categoria",
        "Unidade",
        "Qtd_Atual",
        "Est_Minimo",
        "Est_Maximo",
        "Data_Ultima_Reposicao",
        "Status",
    ],
    "Fornecedores_Precos": [
        "ID",
        "Insumo",
        "Fornecedor",
        "Preco_Unitario",
        "Unidade",
        "Data_Registro",
    ],
    "Entradas_Compras": [
        "ID_Nota",
        "Data",
        "Fornecedor",
        "Valor_Total",
        "Link_Drive",
    ],
    "Saidas_Vendas": [
        "ID_Venda",
        "Data",
        "Prato_Vendido",
        "Qtd_Vendida",
    ],
    "Ficha_Tecnica": [
        "Prato",
        "Insumo",
        "Qtd_Insumo",
        "Unidade",
    ],
}

# Categorias Oficiais do Restaurante Mazon
CATEGORIAS_OFICIAIS = [
    "Proteínas e Frutos do Mar",
    "Hortifruti",
    "Grãos e Mercearia",
    "Alimentação dos Funcionários",
    "Laticínios e Refrigerados",
    "Japoneses e Condimentos",
    "Doces e Confeitaria",
    "Embalagens e Descartáveis",
    "Higiene e Limpeza",
]

SESSION_DB_KEY = "mazon_db_oficial_zerado_v4"


def calcular_status_estoque(qtd_atual: float, est_minimo: float, est_maximo: float = 0.0) -> str:
    """
    Define o status operacional do insumo:
    - PENDENTE: Quando Qtd_Atual, Est_Minimo e Est_Maximo ainda estão zerados (0.0)
    - CRÍTICO/FALTA: Qtd_Atual <= Est_Minimo
    - ALERTA: Qtd_Atual entre Est_Minimo e 25% acima do Est_Minimo
    - OK/CHEIO: Acima da faixa de alerta
    """
    try:
        qtd = float(qtd_atual)
        minimo = float(est_minimo)
        maximo = float(est_maximo)
    except (ValueError, TypeError):
        return "PENDENTE"

    if qtd == 0.0 and minimo == 0.0 and maximo == 0.0:
        return "PENDENTE"
    if qtd <= minimo:
        return "CRÍTICO/FALTA"
    elif qtd <= minimo * 1.25:
        return "ALERTA"
    return "OK/CHEIO"


def _gerar_dados_iniciais() -> Dict[str, pd.DataFrame]:
    """
    Popula todos os 90 insumos oficiais e suas 9 categorias do Restaurante Mazon,
    com valores numéricos zerados para início da operação real:
    - Qtd_Atual: 0.0
    - Est_Minimo: 0.0
    - Est_Maximo: 0.0
    - Data_Ultima_Reposicao: "-"
    - Status: "PENDENTE"
    """
    catalogo_oficial = [
        # 1. Proteínas e Frutos do Mar (9 itens)
        ("Salmão", "Proteínas e Frutos do Mar", "kg"),
        ("Atum", "Proteínas e Frutos do Mar", "kg"),
        ("Peixe branco", "Proteínas e Frutos do Mar", "kg"),
        ("Camarão", "Proteínas e Frutos do Mar", "kg"),
        ("Lula", "Proteínas e Frutos do Mar", "kg"),
        ("Haddock", "Proteínas e Frutos do Mar", "kg"),
        ("Kani", "Proteínas e Frutos do Mar", "pct"),
        ("Filé mignon", "Proteínas e Frutos do Mar", "kg"),
        ("Peito de frango", "Proteínas e Frutos do Mar", "kg"),

        # 2. Hortifruti (17 itens)
        ("Pepino japonês", "Hortifruti", "kg"),
        ("Cebola roxa", "Hortifruti", "kg"),
        ("Tomate", "Hortifruti", "kg"),
        ("Limão", "Hortifruti", "kg"),
        ("Manga", "Hortifruti", "kg"),
        ("Laranja", "Hortifruti", "kg"),
        ("Banana", "Hortifruti", "kg"),
        ("Abacaxi", "Hortifruti", "un"),
        ("Maracujá", "Hortifruti", "kg"),
        ("Repolho verde", "Hortifruti", "un"),
        ("Repolho roxo", "Hortifruti", "un"),
        ("Chuchu", "Hortifruti", "kg"),
        ("Beterraba", "Hortifruti", "kg"),
        ("Cebolinha", "Hortifruti", "maço"),
        ("Coentro", "Hortifruti", "maço"),
        ("Gengibre", "Hortifruti", "kg"),
        ("Pimenta biquinho", "Hortifruti", "kg"),

        # 3. Grãos e Mercearia (15 itens)
        ("Arroz japonês", "Grãos e Mercearia", "kg"),
        ("Farinha panko", "Grãos e Mercearia", "kg"),
        ("Amido de milho", "Grãos e Mercearia", "kg"),
        ("Farinha de trigo", "Grãos e Mercearia", "kg"),
        ("Açúcar cristal", "Grãos e Mercearia", "kg"),
        ("Sal", "Grãos e Mercearia", "kg"),
        ("Macarrão Yakisoba", "Grãos e Mercearia", "pct"),
        ("Ervilha", "Grãos e Mercearia", "kg"),
        ("Gergelim", "Grãos e Mercearia", "kg"),
        ("Floco de arroz", "Grãos e Mercearia", "pct"),
        ("Farinha de mesa", "Grãos e Mercearia", "kg"),
        ("Fubá", "Grãos e Mercearia", "kg"),
        ("Amendoim", "Grãos e Mercearia", "kg"),
        ("Coco ralado", "Grãos e Mercearia", "kg"),
        ("Biscoito Maizena", "Grãos e Mercearia", "pct"),

        # 4. Alimentação dos Funcionários (3 itens)
        ("Arroz funcionário", "Alimentação dos Funcionários", "kg"),
        ("Feijão funcionário", "Alimentação dos Funcionários", "kg"),
        ("Café funcionário", "Alimentação dos Funcionários", "pct"),

        # 5. Laticínios e Refrigerados (6 itens)
        ("Ovo", "Laticínios e Refrigerados", "cartela"),
        ("Creme cheese", "Laticínios e Refrigerados", "kg"),
        ("Catupiry", "Laticínios e Refrigerados", "kg"),
        ("Leite em pó", "Laticínios e Refrigerados", "kg"),
        ("Creme de leite", "Laticínios e Refrigerados", "un"),
        ("Leite condensado", "Laticínios e Refrigerados", "un"),

        # 6. Japoneses e Condimentos (12 itens)
        ("Shoyu", "Japoneses e Condimentos", "L"),
        ("Saque culinário", "Japoneses e Condimentos", "L"),
        ("Nori", "Japoneses e Condimentos", "pct"),
        ("Wassabi", "Japoneses e Condimentos", "pct"),
        ("Tonkatsu", "Japoneses e Condimentos", "L"),
        ("Óleo", "Japoneses e Condimentos", "L"),
        ("Óleo gergelim", "Japoneses e Condimentos", "L"),
        ("Glutamato", "Japoneses e Condimentos", "kg"),
        ("Limão Pepper", "Japoneses e Condimentos", "kg"),
        ("Chimichurri", "Japoneses e Condimentos", "kg"),
        ("Azeitona", "Japoneses e Condimentos", "kg"),
        ("Palmito", "Japoneses e Condimentos", "vidro"),

        # 7. Doces e Confeitaria (6 itens)
        ("Doce de leite", "Doces e Confeitaria", "kg"),
        ("Sococo (Doce cremoso)", "Doces e Confeitaria", "un"),
        ("Calda morango", "Doces e Confeitaria", "un"),
        ("Goiabada", "Doces e Confeitaria", "kg"),
        ("Goiabada líquida", "Doces e Confeitaria", "kg"),
        ("Barra de chocolate", "Doces e Confeitaria", "kg"),

        # 8. Embalagens e Descartáveis (12 itens)
        ("H01", "Embalagens e Descartáveis", "un"),
        ("H02", "Embalagens e Descartáveis", "un"),
        ("H03", "Embalagens e Descartáveis", "un"),
        ("S640 (Pote tempura de banana)", "Embalagens e Descartáveis", "un"),
        ("Pote molho (30 ml)", "Embalagens e Descartáveis", "un"),
        ("Saco Kraft P", "Embalagens e Descartáveis", "un"),
        ("Saco Kraft M", "Embalagens e Descartáveis", "un"),
        ("Saco Kraft G", "Embalagens e Descartáveis", "un"),
        ("Saco Kraft GG", "Embalagens e Descartáveis", "un"),
        ("Hashi", "Embalagens e Descartáveis", "par"),
        ("Nozoki", "Embalagens e Descartáveis", "un"),
        ("Palito", "Embalagens e Descartáveis", "cx"),

        # 9. Higiene e Limpeza (10 itens)
        ("Papel toalha", "Higiene e Limpeza", "pct"),
        ("Touca", "Higiene e Limpeza", "cx"),
        ("Gás (Maçarico)", "Higiene e Limpeza", "un"),
        ("Sbp", "Higiene e Limpeza", "un"),
        ("Álcool", "Higiene e Limpeza", "L"),
        ("Detergente", "Higiene e Limpeza", "L"),
        ("Perfex", "Higiene e Limpeza", "rolo"),
        ("Papel higiênico", "Higiene e Limpeza", "fardo"),
        ("Saco de lixo", "Higiene e Limpeza", "pct"),
        ("Desinfetante", "Higiene e Limpeza", "L"),
    ]

    estoque_dicts = []
    for idx, (nome, cat, unid) in enumerate(catalogo_oficial, start=1):
        estoque_dicts.append(
            {
                "ID": f"MZ-{idx:03d}",
                "Item": nome,
                "Categoria": cat,
                "Unidade": unid,
                "Qtd_Atual": 0.0,
                "Est_Minimo": 0.0,
                "Est_Maximo": 0.0,
                "Data_Ultima_Reposicao": "-",
                "Status": "PENDENTE",
            }
        )

    df_estoque = pd.DataFrame(estoque_dicts, columns=SHEETS_SCHEMA["Estoque"])
    df_precos = pd.DataFrame(columns=SHEETS_SCHEMA["Fornecedores_Precos"])
    df_entradas = pd.DataFrame(columns=SHEETS_SCHEMA["Entradas_Compras"])
    df_saidas = pd.DataFrame(columns=SHEETS_SCHEMA["Saidas_Vendas"])
    df_ficha = pd.DataFrame(columns=SHEETS_SCHEMA["Ficha_Tecnica"])

    return {
        "Estoque": df_estoque,
        "Fornecedores_Precos": df_precos,
        "Entradas_Compras": df_entradas,
        "Saidas_Vendas": df_saidas,
        "Ficha_Tecnica": df_ficha,
    }


class MazonSheetsManager:
    """
    Gerenciador de persistência híbrida (Google Sheets + Google Drive + Fallback em Sessão).
    Garante que o sistema opere conectado à nuvem ou em modo local com o estoque oficial zerado.
    """

    def __init__(
        self,
        credentials_path: str = "credentials.json",
        spreadsheet_name: str = "Restaurante_Mazon_ERP_Database",
        spreadsheet_id: Optional[str] = None,
    ) -> None:
        self.credentials_path = credentials_path
        self.spreadsheet_name = spreadsheet_name
        self.spreadsheet_id = spreadsheet_id or os.getenv("MAZON_SPREADSHEET_ID")
        self.client: Optional[gspread.Client] = None
        self.spreadsheet: Optional[gspread.Spreadsheet] = None
        self.drive_service: Any = None
        self.connected: bool = False
        self.connection_message: str = "Base Oficial Mazon Ativa (90 insumos zerados)"

        self._init_local_state()
        self._try_connect_cloud()

    def _init_local_state(self) -> None:
        """
        Força a limpeza do cache antigo de memória na primeira execução da nova versão
        e inicializa os DataFrames com os valores zerados oficiais.
        """
        if "mazon_cache_reset_v3" not in st.session_state:
            st.session_state.estoque_df = None
            for chave_antiga in ["mazon_db", "mazon_db_oficial_v2"]:
                st.session_state.pop(chave_antiga, None)
            st.session_state[SESSION_DB_KEY] = _gerar_dados_iniciais()
            st.session_state["mazon_cache_reset_v3"] = True
        elif SESSION_DB_KEY not in st.session_state:
            st.session_state[SESSION_DB_KEY] = _gerar_dados_iniciais()

    def _try_connect_cloud(self) -> None:
        """Tenta autenticar no Google Cloud via Service Account ou st.secrets."""
        creds: Optional[Credentials] = None

        try:
            if os.path.exists(self.credentials_path):
                creds = Credentials.from_service_account_file(
                    self.credentials_path, scopes=SCOPES
                )
            elif "gcp_service_account" in st.secrets:
                creds = Credentials.from_service_account_info(
                    dict(st.secrets["gcp_service_account"]), scopes=SCOPES
                )
        except Exception as exc:
            self.connection_message = f"Credenciais não carregadas ({exc})"
            self.connected = False
            return

        if creds is None:
            self.connection_message = (
                "Base Oficial Zerada Ativa (Adicione 'credentials.json' para sincronizar na nuvem)."
            )
            self.connected = False
            return

        try:
            self.client = gspread.authorize(creds)
            self.drive_service = build("drive", "v3", credentials=creds)

            if self.spreadsheet_id:
                self.spreadsheet = self.client.open_by_key(self.spreadsheet_id)
            else:
                try:
                    self.spreadsheet = self.client.open(self.spreadsheet_name)
                except gspread.SpreadsheetNotFound:
                    self.spreadsheet = self.client.create(self.spreadsheet_name)

            self._garantir_abas_obrigatorias()
            self.connected = True
            self.connection_message = f"Conectado ao Google Sheets: {self.spreadsheet.title}"
        except Exception as exc:
            self.connected = False
            self.connection_message = f"Modo Local Ativo (Aviso Google Cloud: {exc})"

    def _garantir_abas_obrigatorias(self) -> None:
        """Verifica e cria automaticamente as 5 abas obrigatórias no Google Sheets com seus cabeçalhos."""
        if not self.spreadsheet:
            return

        abas_existentes = {ws.title: ws for ws in self.spreadsheet.worksheets()}
        dados_iniciais = _gerar_dados_iniciais()

        for nome_aba, colunas in SHEETS_SCHEMA.items():
            if nome_aba not in abas_existentes:
                ws = self.spreadsheet.add_worksheet(title=nome_aba, rows=500, cols=len(colunas) + 2)
                df_seed = dados_iniciais[nome_aba]
                valores = [colunas] + df_seed.astype(str).values.tolist()
                ws.update(valores, value_input_option="USER_ENTERED")
            else:
                ws = abas_existentes[nome_aba]
                primeira_linha = ws.row_values(1)
                if not primeira_linha:
                    df_seed = dados_iniciais[nome_aba]
                    valores = [colunas] + df_seed.astype(str).values.tolist()
                    ws.update(valores, value_input_option="USER_ENTERED")

    def ler_aba(self, nome_aba: str) -> pd.DataFrame:
        """Lê uma das 5 abas do Google Sheets (ou do estado local se offline)."""
        colunas_esperadas = SHEETS_SCHEMA[nome_aba]

        if self.connected and self.spreadsheet:
            try:
                ws = self.spreadsheet.worksheet(nome_aba)
                registros = ws.get_all_records()
                if not registros:
                    return pd.DataFrame(columns=colunas_esperadas)
                df = pd.DataFrame(registros)
                for col in colunas_esperadas:
                    if col not in df.columns:
                        df[col] = ""
                df = df[colunas_esperadas]
                st.session_state[SESSION_DB_KEY][nome_aba] = df.copy()
                return df
            except Exception:
                pass

        return st.session_state[SESSION_DB_KEY][nome_aba].copy()

    def salvar_aba(self, nome_aba: str, df: pd.DataFrame) -> bool:
        """Sobrescreve a aba de forma segura no estado local e sincroniza com o Google Sheets."""
        colunas = SHEETS_SCHEMA[nome_aba]
        df_limpo = df[colunas].copy()
        st.session_state[SESSION_DB_KEY][nome_aba] = df_limpo

        if self.connected and self.spreadsheet:
            try:
                ws = self.spreadsheet.worksheet(nome_aba)
                ws.clear()
                valores = [colunas] + df_limpo.astype(str).values.tolist()
                ws.update(valores, value_input_option="USER_ENTERED")
                return True
            except Exception as exc:
                st.warning(f"Salvo localmente, mas houve falha ao sincronizar no Sheets: {exc}")
                return False
        return True

    def obter_estoque_enriquecido(self) -> pd.DataFrame:
        """
        Retorna a tabela 'Estoque' enriquecida com o último Preço Unitário de 'Fornecedores_Precos'
        e o Valor Total em Estoque (R$) de cada item.
        """
        df_estoque = self.ler_aba("Estoque").copy()
        df_precos = self.ler_aba("Fornecedores_Precos").copy()

        for col_num in ["Qtd_Atual", "Est_Minimo", "Est_Maximo"]:
            df_estoque[col_num] = pd.to_numeric(df_estoque[col_num], errors="coerce").fillna(0.0)

        df_estoque["Status"] = df_estoque.apply(
            lambda r: calcular_status_estoque(r["Qtd_Atual"], r["Est_Minimo"], r["Est_Maximo"]),
            axis=1,
        )

        if not df_precos.empty:
            df_precos["Preco_Unitario"] = pd.to_numeric(
                df_precos["Preco_Unitario"], errors="coerce"
            ).fillna(0.0)
            df_precos_ord = df_precos.sort_values("Data_Registro", ascending=False)
            ultimos_precos = (
                df_precos_ord.groupby("Insumo", as_index=False)["Preco_Unitario"]
                .first()
                .rename(columns={"Insumo": "Item", "Preco_Unitario": "Ultimo_Preco_Unitario"})
            )
            df_estoque = df_estoque.merge(ultimos_precos, on="Item", how="left")
            df_estoque["Ultimo_Preco_Unitario"] = df_estoque["Ultimo_Preco_Unitario"].fillna(0.0)
        else:
            df_estoque["Ultimo_Preco_Unitario"] = 0.0

        df_estoque["Valor_Total_Item"] = (
            df_estoque["Qtd_Atual"] * df_estoque["Ultimo_Preco_Unitario"]
        )
        return df_estoque

    def cadastrar_ou_atualizar_item_estoque(
        self,
        item_id: str,
        nome_item: str,
        categoria: str,
        unidade: str,
        qtd_atual: float,
        est_minimo: float,
        est_maximo: float,
        preco_referencia: Optional[float] = None,
        fornecedor_referencia: str = "Cadastro Direto",
    ) -> Tuple[bool, str]:
        """Cadastra novo insumo ou atualiza um existente sincronizando com o Google Sheets."""
        df_estoque = self.ler_aba("Estoque")
        hoje = datetime.now().strftime("%Y-%m-%d")
        novo_status = calcular_status_estoque(qtd_atual, est_minimo, est_maximo)

        if not item_id.strip():
            proximo_num = len(df_estoque) + 1
            item_id = f"MZ-{proximo_num:03d}"

        mascara = (df_estoque["ID"].astype(str) == item_id) | (
            df_estoque["Item"].astype(str).str.lower() == nome_item.strip().lower()
        )

        if mascara.any():
            idx = df_estoque[mascara].index[0]
            df_estoque.loc[idx, "Item"] = nome_item.strip()
            df_estoque.loc[idx, "Categoria"] = categoria
            df_estoque.loc[idx, "Unidade"] = unidade
            df_estoque.loc[idx, "Qtd_Atual"] = float(qtd_atual)
            df_estoque.loc[idx, "Est_Minimo"] = float(est_minimo)
            df_estoque.loc[idx, "Est_Maximo"] = float(est_maximo)
            df_estoque.loc[idx, "Data_Ultima_Reposicao"] = hoje
            df_estoque.loc[idx, "Status"] = novo_status
            acao_msg = f"Insumo '{nome_item}' ({df_estoque.loc[idx, 'ID']}) atualizado com sucesso!"
        else:
            nova_linha = pd.DataFrame(
                [
                    {
                        "ID": item_id,
                        "Item": nome_item.strip(),
                        "Categoria": categoria,
                        "Unidade": unidade,
                        "Qtd_Atual": float(qtd_atual),
                        "Est_Minimo": float(est_minimo),
                        "Est_Maximo": float(est_maximo),
                        "Data_Ultima_Reposicao": hoje,
                        "Status": novo_status,
                    }
                ]
            )
            df_estoque = pd.concat([df_estoque, nova_linha], ignore_index=True)
            acao_msg = f"Novo insumo '{nome_item}' ({item_id}) cadastrado com sucesso!"

        ok = self.salvar_aba("Estoque", df_estoque)

        if preco_referencia and preco_referencia > 0:
            self.registrar_preco_fornecedor(
                insumo=nome_item.strip(),
                fornecedor=fornecedor_referencia,
                preco_unitario=preco_referencia,
                unidade=unidade,
                data_registro=hoje,
            )

        return ok, acao_msg

    def registrar_preco_fornecedor(
        self,
        insumo: str,
        fornecedor: str,
        preco_unitario: float,
        unidade: str,
        data_registro: Optional[str] = None,
    ) -> bool:
        """Adiciona cotação ou preço de compra na aba 'Fornecedores_Precos'."""
        df_precos = self.ler_aba("Fornecedores_Precos")
        novo_id = f"FP-{len(df_precos) + 1:03d}"
        data_str = data_registro or datetime.now().strftime("%Y-%m-%d")

        nova_cotacao = pd.DataFrame(
            [
                {
                    "ID": novo_id,
                    "Insumo": insumo.strip(),
                    "Fornecedor": fornecedor.strip(),
                    "Preco_Unitario": round(float(preco_unitario), 2),
                    "Unidade": unidade.strip(),
                    "Data_Registro": data_str,
                }
            ]
        )
        df_precos = pd.concat([df_precos, nova_cotacao], ignore_index=True)
        return self.salvar_aba("Fornecedores_Precos", df_precos)

    def upload_arquivo_drive(
        self, arquivo_bytes: bytes, nome_arquivo: str, mime_type: str, folder_id: Optional[str] = None
    ) -> str:
        """
        Envia PDF/Imagem de Nota Fiscal ou Extrato para o Google Drive.
        Retorna o link compartilhável (webViewLink) ou URI simulada em modo Demo.
        """
        if self.connected and self.drive_service:
            try:
                metadata: Dict[str, Any] = {"name": nome_arquivo}
                if folder_id:
                    metadata["parents"] = [folder_id]
                media = MediaIoBaseUpload(
                    io.BytesIO(arquivo_bytes), mimetype=mime_type, resumable=False
                )
                created = (
                    self.drive_service.files()
                    .create(body=metadata, media_body=media, fields="id, webViewLink")
                    .execute()
                )
                return created.get(
                    "webViewLink", f"https://drive.google.com/file/d/{created.get('id')}"
                )
            except Exception as exc:
                return f"https://drive.google.com/local-fallback/{nome_arquivo} (Aviso: {exc})"

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"https://drive.google.com/drive/u/0/mazon-notas/{timestamp}_{nome_arquivo}"


# ==============================================================================
# ESTRUTURA PREPARADA PARA FASE 2: OCR / VISÃO COMPUTACIONAL (GEMINI / CLOUD VISION)
# ==============================================================================
def processar_documento_ocr_ia(
    arquivo_bytes: bytes,
    nome_arquivo: str,
    tipo_documento: str,
) -> Dict[str, Any]:
    """
    Hook estruturado para a Fase 2 (Integração Google Cloud Vision / Gemini Flash Vision).
    Recebe os bytes da imagem/PDF, executa o pipeline de pré-processamento e retorna
    o payload estruturado para aprovação e gravação nas abas 'Entradas_Compras' ou 'Saidas_Vendas'.
    """
    tamanho_kb = round(len(arquivo_bytes) / 1024, 2)
    agora = datetime.now().strftime("%Y-%m-%d %H:%M")

    if tipo_documento == "Nota Fiscal / Cupom de Compra":
        return {
            "status_pipeline": "PRONTO_PARA_FASE_2_OCR",
            "documento": nome_arquivo,
            "tamanho_kb": tamanho_kb,
            "processado_em": agora,
            "aba_destino": "Entradas_Compras + Estoque + Fornecedores_Precos",
            "campos_extraidos_preview": {
                "ID_Nota": f"NF-AUTO-{datetime.now().strftime('%H%M')}",
                "Data": datetime.now().strftime("%Y-%m-%d"),
                "Fornecedor": "Aguardando leitura OCR Fase 2",
                "Valor_Total": 0.0,
                "Itens_Detectados": [],
            },
        }

    return {
        "status_pipeline": "PRONTO_PARA_FASE_2_OCR",
        "documento": nome_arquivo,
        "tamanho_kb": tamanho_kb,
        "processado_em": agora,
        "aba_destino": "Saidas_Vendas + Baixa via Ficha_Tecnica",
        "campos_extraidos_preview": {
            "ID_Venda": f"VD-AUTO-{datetime.now().strftime('%H%M')}",
            "Data": datetime.now().strftime("%Y-%m-%d"),
            "Pratos_Extraidos": [],
        },
    }
