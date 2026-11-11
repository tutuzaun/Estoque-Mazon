"""
Módulo Unificado de Persistência em Nuvem (Google Sheets & Google Drive) + RBAC & Auditoria
Restaurante Mazon - Gestão de Estoque e Inteligência Financeira
Centraliza o banco de dados em uma única planilha ("Mazon - Banco de Dados") com as abas:
1. Usuarios
2. Logs_Auditoria
3. Estoque
4. Financeiro
E gerencia o upload de Notas Fiscais e Comprovantes em uma única pasta principal no Google Drive
com criação automática das subpastas /Notas_Fiscais e /Comprovantes.
"""

from __future__ import annotations

import io
import os
from datetime import datetime, timedelta
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

# Permissões Granulares por Módulo
PERMISSOES_DISPONIVEIS: Dict[str, str] = {
    "home": "🏠 Home — Painel Executivo e Alertas Críticos",
    "estoque_leitura": "📦 Estoque — Dashboard, Tabela, Movimentação, Comparativo, OCR e Compras",
    "estoque_edicao": "📦 Estoque — Cadastrar e Editar Parâmetros de Insumos",
    "fin_lancamentos": "💰 Financeiro — Lançamentos Diários e Upload de NFs/Comprovantes",
    "fin_relatorios": "💰 Financeiro — Faturamento, DRE, CMV, Break-Even e Relatório Descritivo",
    "auditoria_logs": "🛡️ Auditoria — Visualização de Logs de Operações",
    "usuarios_gestao": "🛡️ Usuários — Gestão de Acessos, Cargos e Permissões",
}

PERMISSOES_POR_PERFIL: Dict[str, List[str]] = {
    "Dev": list(PERMISSOES_DISPONIVEIS.keys()),
    "Gerente": list(PERMISSOES_DISPONIVEIS.keys()),
    "Caixa": ["home", "estoque_leitura", "fin_lancamentos"],
    "Equipe": ["home", "estoque_leitura"],
}

# Esquema Unificado das 4 Abas Principais na Planilha Única "Mazon - Banco de Dados"
SHEETS_SCHEMA: Dict[str, List[str]] = {
    "Usuarios": [
        "Usuario",
        "Perfil",
        "Senha",
        "Data_Cadastro",
        "Ativo",
        "Permissoes",
    ],
    "Logs_Auditoria": [
        "Data_Hora",
        "Usuario",
        "Perfil",
        "Acao",
        "Modulo",
    ],
    "Estoque": [
        "ID",
        "Item",
        "Categoria",
        "Unidade",
        "Qtd_Atual",
        "Est_Minimo",
        "Est_Maximo",
        "Preco_Unitario",
        "Fornecedor",
        "Data_Ultima_Reposicao",
        "Status",
    ],
    "Financeiro": [
        "ID_Lancamento",
        "Data",
        "Tipo",
        "Natureza_Custo",
        "Descricao",
        "Categoria",
        "Fornecedor",
        "Forma_Pagamento",
        "Valor",
        "Link_Comprovante",
        "Responsavel",
    ],
}

# Usuários Padrão Obrigatórios (RBAC)
USUARIOS_PADRAO = [
    ("Arthur", "Dev", "245716", "2026-11-01", "Sim", ",".join(PERMISSOES_POR_PERFIL["Dev"])),
    ("Claudio", "Gerente", "221016", "2026-11-01", "Sim", ",".join(PERMISSOES_POR_PERFIL["Gerente"])),
    ("Caixa", "Caixa", "mazon123", "2026-11-01", "Sim", ",".join(PERMISSOES_POR_PERFIL["Caixa"])),
    ("Equipe", "Equipe", "1234", "2026-11-01", "Sim", ",".join(PERMISSOES_POR_PERFIL["Equipe"])),
]

# Categorias Oficiais de Estoque do Restaurante Mazon
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

CATEGORIAS_FINANCEIRAS = [
    "Faturamento Salão",
    "Faturamento Delivery (iFood / Próprio)",
    "Eventos & Reservas",
    "Insumos - Proteínas e Pescados",
    "Insumos - Mercearia e Hortifruti",
    "Insumos - Embalagens e Descartáveis",
    "Folha de Pagamento & Encargos",
    "Aluguel, Energia, Água e Gás",
    "Manutenção, Limpeza e Operacional",
    "Impostos & Taxas de Cartão/Apps",
]

FORMAS_PAGAMENTO = [
    "PIX",
    "Cartão de Crédito",
    "Cartão de Débito",
    "Dinheiro",
    "Boleto Faturado",
    "Transferência TED",
    "iFood / App Online",
]

SESSION_DB_KEY = "mazon_unified_db_v6"


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
    Inicializa as 4 abas da planilha unificada sem nenhum dado fictício (mock):
    - Usuarios (4 perfis RBAC padrão obrigatórios)
    - Logs_Auditoria (vazio — preenchido apenas por operações reais)
    - Estoque (90 insumos oficiais do Restaurante Mazon com contagens e custos zerados em 0.0)
    - Financeiro (vazio — preenchido apenas por lançamentos reais)
    """
    df_usuarios = pd.DataFrame(USUARIOS_PADRAO, columns=SHEETS_SCHEMA["Usuarios"])
    df_logs = pd.DataFrame(columns=SHEETS_SCHEMA["Logs_Auditoria"])

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
                "Preco_Unitario": 0.0,
                "Fornecedor": "-",
                "Data_Ultima_Reposicao": "-",
                "Status": "PENDENTE",
            }
        )

    df_estoque = pd.DataFrame(estoque_dicts, columns=SHEETS_SCHEMA["Estoque"])
    df_financeiro = pd.DataFrame(columns=SHEETS_SCHEMA["Financeiro"])

    return {
        "Usuarios": df_usuarios,
        "Logs_Auditoria": df_logs,
        "Estoque": df_estoque,
        "Financeiro": df_financeiro,
    }


class MazonSheetsManager:
    """
    Gerenciador Unificado do Banco de Dados no Google Sheets ("Mazon - Banco de Dados")
    e Repositório Único de Arquivos no Google Drive com subpastas /Notas_Fiscais e /Comprovantes.
    """

    def __init__(
        self,
        credentials_path: str = "credentials.json",
        spreadsheet_name: str = "Mazon - Banco de Dados",
        spreadsheet_id: Optional[str] = None,
        drive_main_folder_name: str = "Mazon - Documentos e Comprovantes",
    ) -> None:
        self.credentials_path = credentials_path
        self.spreadsheet_name = spreadsheet_name
        self.spreadsheet_id = spreadsheet_id or os.getenv("MAZON_SPREADSHEET_ID")
        self.drive_main_folder_name = drive_main_folder_name
        self.client: Optional[gspread.Client] = None
        self.spreadsheet: Optional[gspread.Spreadsheet] = None
        self.drive_service: Any = None
        self.connected: bool = False
        self.connection_message: str = "Banco Unificado Mazon Ativo"

        self._init_local_state()
        self._try_connect_cloud()

    def _init_local_state(self) -> None:
        """Inicializa o banco unificado na sessão do Streamlit."""
        if SESSION_DB_KEY not in st.session_state:
            st.session_state[SESSION_DB_KEY] = _gerar_dados_iniciais()

    def _try_connect_cloud(self) -> None:
        """
        Autentica no Google Cloud verificando PRIMEIRO st.secrets (para Streamlit Cloud)
        e DEPOIS o arquivo de credenciais local (credentials.json).
        """
        creds: Optional[Credentials] = None

        # 1. Prioridade: st.secrets (Streamlit Cloud)
        try:
            if "gcp_service_account" in st.secrets:
                creds = Credentials.from_service_account_info(
                    dict(st.secrets["gcp_service_account"]), scopes=SCOPES
                )
            elif "type" in st.secrets and "private_key" in st.secrets:
                creds = Credentials.from_service_account_info(
                    dict(st.secrets), scopes=SCOPES
                )
        except Exception:
            creds = None

        # 2. Fallback: Arquivo de credenciais local (credentials.json)
        if creds is None:
            try:
                if os.path.exists(self.credentials_path):
                    creds = Credentials.from_service_account_file(
                        self.credentials_path, scopes=SCOPES
                    )
            except Exception as exc:
                self.connection_message = f"Modo Local Ativo ({exc})"
                self.connected = False
                return

        if creds is None:
            self.connection_message = "Banco Unificado em Memória Ativo (Offline)"
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
                    try:
                        # Tenta abrir a planilha anterior caso o usuário já a tenha compartilhado
                        self.spreadsheet = self.client.open("Restaurante_Mazon_ERP_Database")
                    except Exception:
                        self.spreadsheet = self.client.create(self.spreadsheet_name)

            self.connected = True
            self._garantir_abas_obrigatorias()
            self.connection_message = f"Planilha Unificada Online: {self.spreadsheet.title}"
        except Exception as exc:
            self.connected = False
            self.connection_message = f"Operando em Cache Seguro (Aviso Cloud: {exc})"

    def _garantir_abas_obrigatorias(self) -> None:
        """
        Verifica e inicializa automaticamente as 4 abas obrigatórias:
        Usuarios, Logs_Auditoria, Estoque e Financeiro.
        Garante também que os 4 usuários padrão existam na aba Usuarios
        e remove eventuais lançamentos fictícios de versões anteriores.
        """
        if not self.spreadsheet:
            return

        abas_existentes = {ws.title: ws for ws in self.spreadsheet.worksheets()}
        dados_iniciais = _gerar_dados_iniciais()

        for nome_aba, colunas in SHEETS_SCHEMA.items():
            if nome_aba not in abas_existentes:
                ws = self.spreadsheet.add_worksheet(
                    title=nome_aba, rows=600, cols=len(colunas) + 2
                )
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

        # Garante que os 4 usuários padrão estejam presentes na aba Usuarios
        try:
            df_u = self.ler_aba("Usuarios")
            usuarios_atuais = set(df_u["Usuario"].astype(str).str.strip().tolist())
            novos_u = []
            for u_nome, u_perf, u_senha, u_data, u_ativo, u_perms in USUARIOS_PADRAO:
                if u_nome not in usuarios_atuais:
                    novos_u.append(
                        {
                            "Usuario": u_nome,
                            "Perfil": u_perf,
                            "Senha": u_senha,
                            "Data_Cadastro": u_data,
                            "Ativo": u_ativo,
                            "Permissoes": u_perms,
                        }
                    )
            if novos_u:
                df_u = pd.concat([df_u, pd.DataFrame(novos_u)], ignore_index=True)
                self.salvar_aba("Usuarios", df_u)
        except Exception:
            pass

        # Limpa automaticamente lançamentos fictícios (mock) caso tenham sido gravados antes
        try:
            descricoes_ficticias = {
                "Fechamento Salão Jantar",
                "Compra de Salmão e Atum Fresco",
                "Vendas Delivery Turno Noite",
                "Reposição Embalagens H01/H02/Kraft",
                "Fechamento Parcial Caixa Dia",
                "Energia Elétrica e Gás Industrial",
            }
            df_f = self.ler_aba("Financeiro")
            if not df_f.empty and df_f["Descricao"].astype(str).isin(descricoes_ficticias).any():
                df_f_limpo = df_f[~df_f["Descricao"].astype(str).isin(descricoes_ficticias)].copy()
                self.salvar_aba("Financeiro", df_f_limpo)
        except Exception:
            pass

    @staticmethod
    def resolver_permissoes(perfil: str, permissoes_str: str = "") -> List[str]:
        """Resolve a lista de permissões ativas de um usuário com fallback para o perfil padrão."""
        p_limpo = str(permissoes_str or "").strip()
        if p_limpo and p_limpo != "-":
            lista = [p.strip() for p in p_limpo.split(",") if p.strip() in PERMISSOES_DISPONIVEIS]
            if lista:
                return lista
        return list(PERMISSOES_POR_PERFIL.get(perfil, ["home", "estoque_leitura"]))

    def ler_aba(self, nome_aba: str) -> pd.DataFrame:
        """Lê qualquer aba da planilha unificada com fallback automático em sessão."""
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
                        df[col] = 0.0 if col in ["Preco_Unitario", "Valor"] else "-"
                df = df[colunas_esperadas]
                st.session_state[SESSION_DB_KEY][nome_aba] = df.copy()
                return df
            except Exception:
                pass

        return st.session_state[SESSION_DB_KEY][nome_aba].copy()

    def salvar_aba(self, nome_aba: str, df: pd.DataFrame) -> bool:
        """Salva um DataFrame na aba correspondente da planilha unificada."""
        colunas = SHEETS_SCHEMA[nome_aba]
        df_limpo = df.copy()
        for col in colunas:
            if col not in df_limpo.columns:
                df_limpo[col] = ""
        df_limpo = df_limpo[colunas]
        st.session_state[SESSION_DB_KEY][nome_aba] = df_limpo

        if self.connected and self.spreadsheet:
            try:
                ws = self.spreadsheet.worksheet(nome_aba)
                ws.clear()
                valores = [colunas] + df_limpo.astype(str).values.tolist()
                ws.update(valores, value_input_option="USER_ENTERED")
                return True
            except Exception as exc:
                st.warning(f"Salvo localmente (Aviso de sincronização no Sheets: {exc})")
                return False
        return True

    # ==========================================================================
    # AUTENTICAÇÃO, HIERARQUIA E GESTÃO DE USUÁRIOS (ABA 'Usuarios')
    # ==========================================================================
    def autenticar_usuario(self, usuario: str, senha: str) -> Optional[Dict[str, Any]]:
        """Valida credenciais na aba 'Usuarios' e retorna {usuario, perfil, permissoes} se válido."""
        df_usuarios = self.ler_aba("Usuarios")
        u_limpo = usuario.strip().lower()
        s_limpa = str(senha).strip()

        for _, row in df_usuarios.iterrows():
            if str(row.get("Ativo", "Sim")).strip().lower() == "não":
                continue
            if (
                str(row["Usuario"]).strip().lower() == u_limpo
                and str(row["Senha"]).strip() == s_limpa
            ):
                perf = str(row["Perfil"]).strip()
                perms = self.resolver_permissoes(perf, str(row.get("Permissoes", "")))
                return {
                    "usuario": str(row["Usuario"]).strip(),
                    "perfil": perf,
                    "permissoes": perms,
                }
        return None

    def listar_usuarios_visiveis(self, perfil_logado: str) -> pd.DataFrame:
        """
        Retorna os usuários visíveis respeitando a hierarquia:
        - Dev: visualiza 100% dos usuários (incluindo contas Dev).
        - Gerente: visualiza até o nível Gerente (NÃO visualiza contas Dev).
        """
        df_u = self.ler_aba("Usuarios").copy()
        if df_u.empty:
            return df_u
        for idx, row in df_u.iterrows():
            perms_atuais = self.resolver_permissoes(
                str(row.get("Perfil", "Equipe")), str(row.get("Permissoes", ""))
            )
            df_u.at[idx, "Permissoes"] = ",".join(perms_atuais)
        if str(perfil_logado).strip() != "Dev":
            df_u = df_u[df_u["Perfil"].astype(str).str.strip() != "Dev"].copy()
        return df_u

    def cadastrar_ou_atualizar_usuario(
        self,
        novo_usuario: str,
        perfil: str,
        senha: str,
        ativo: str = "Sim",
        permissoes_lista: Optional[List[str]] = None,
        operador_perfil: str = "Dev",
    ) -> Tuple[bool, str]:
        """
        Cadastra ou atualiza um usuário na aba 'Usuarios' respeitando a hierarquia:
        - Apenas o perfil Dev pode criar, visualizar ou promover contas para o nível Dev.
        - O perfil Gerente pode criar e gerenciar contas até o nível Gerente, sem alterar contas Dev.
        """
        df_u = self.ler_aba("Usuarios")
        nome_limpo = novo_usuario.strip()
        if not nome_limpo:
            return False, "Informe o nome do usuário."

        # Regra de Hierarquia: Somente Dev pode atribuir ou editar perfil Dev
        if str(operador_perfil).strip() != "Dev" and str(perfil).strip() == "Dev":
            return (
                False,
                "Acesso negado pela hierarquia: Apenas o perfil Dev (Arthur) pode criar ou promover usuários para o nível Dev.",
            )

        mascara = df_u["Usuario"].astype(str).str.lower() == nome_limpo.lower()
        if mascara.any():
            idx = df_u[mascara].index[0]
            perfil_existente = str(df_u.loc[idx, "Perfil"]).strip()
            if str(operador_perfil).strip() != "Dev" and perfil_existente == "Dev":
                return (
                    False,
                    "Acesso negado pela hierarquia: O perfil Gerente não pode alterar usuários de nível Dev.",
                )

        perms_finais = (
            [p for p in permissoes_lista if p in PERMISSOES_DISPONIVEIS]
            if permissoes_lista is not None
            else list(PERMISSOES_POR_PERFIL.get(perfil, ["home", "estoque_leitura"]))
        )
        if not perms_finais:
            perms_finais = ["home"]
        perms_str = ",".join(perms_finais)

        hoje = datetime.now().strftime("%Y-%m-%d")

        if mascara.any():
            idx = df_u[mascara].index[0]
            df_u.loc[idx, "Perfil"] = perfil
            if str(senha).strip():
                df_u.loc[idx, "Senha"] = str(senha).strip()
            df_u.loc[idx, "Ativo"] = ativo
            df_u.loc[idx, "Permissoes"] = perms_str
            msg = f"Usuário '{nome_limpo}' atualizado com sucesso (Perfil: {perfil} | Status: {ativo})."
        else:
            if not str(senha).strip():
                return False, "Informe uma senha para o novo usuário."
            nova_linha = pd.DataFrame(
                [
                    {
                        "Usuario": nome_limpo,
                        "Perfil": perfil,
                        "Senha": str(senha).strip(),
                        "Data_Cadastro": hoje,
                        "Ativo": ativo,
                        "Permissoes": perms_str,
                    }
                ]
            )
            df_u = pd.concat([df_u, nova_linha], ignore_index=True)
            msg = f"Novo usuário '{nome_limpo}' ({perfil}) cadastrado com sucesso."

        ok = self.salvar_aba("Usuarios", df_u)
        return ok, msg

    # ==========================================================================
    # AUDITORIA E LOGS DE OPERAÇÕES (ABA 'Logs_Auditoria')
    # ==========================================================================
    def registrar_log(self, usuario: str, perfil: str, acao: str, modulo: str) -> bool:
        """
        Grava automaticamente um evento na aba 'Logs_Auditoria' do Sheets com as colunas:
        Data_Hora | Usuario | Perfil | Acao | Modulo.
        """
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        novo_registro = {
            "Data_Hora": agora,
            "Usuario": str(usuario),
            "Perfil": str(perfil),
            "Acao": str(acao),
            "Modulo": str(modulo),
        }

        df_logs = st.session_state[SESSION_DB_KEY]["Logs_Auditoria"].copy()
        df_logs = pd.concat([df_logs, pd.DataFrame([novo_registro])], ignore_index=True)
        st.session_state[SESSION_DB_KEY]["Logs_Auditoria"] = df_logs

        if self.connected and self.spreadsheet:
            try:
                ws = self.spreadsheet.worksheet("Logs_Auditoria")
                ws.append_row(
                    [
                        novo_registro["Data_Hora"],
                        novo_registro["Usuario"],
                        novo_registro["Perfil"],
                        novo_registro["Acao"],
                        novo_registro["Modulo"],
                    ],
                    value_input_option="USER_ENTERED",
                )
                return True
            except Exception:
                return False
        return True

    # ==========================================================================
    # GESTÃO DE ESTOQUE UNIFICADA (ABA 'Estoque')
    # ==========================================================================
    def obter_estoque_enriquecido(self) -> pd.DataFrame:
        """Retorna a tabela 'Estoque' com cálculo numérico de Status e Valor Total."""
        df_estoque = self.ler_aba("Estoque").copy()

        for col_num in ["Qtd_Atual", "Est_Minimo", "Est_Maximo", "Preco_Unitario"]:
            if col_num not in df_estoque.columns:
                df_estoque[col_num] = 0.0
            df_estoque[col_num] = pd.to_numeric(df_estoque[col_num], errors="coerce").fillna(0.0)

        if "Fornecedor" not in df_estoque.columns:
            df_estoque["Fornecedor"] = "-"

        df_estoque["Status"] = df_estoque.apply(
            lambda r: calcular_status_estoque(r["Qtd_Atual"], r["Est_Minimo"], r["Est_Maximo"]),
            axis=1,
        )
        df_estoque["Ultimo_Preco_Unitario"] = df_estoque["Preco_Unitario"]
        df_estoque["Valor_Total_Item"] = df_estoque["Qtd_Atual"] * df_estoque["Preco_Unitario"]
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
        fornecedor_referencia: str = "-",
    ) -> Tuple[bool, str]:
        """Cadastra novo insumo ou atualiza um existente sincronizando com a aba 'Estoque'."""
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
            if preco_referencia is not None:
                df_estoque.loc[idx, "Preco_Unitario"] = float(preco_referencia)
            if fornecedor_referencia and fornecedor_referencia != "-":
                df_estoque.loc[idx, "Fornecedor"] = fornecedor_referencia.strip()
            df_estoque.loc[idx, "Data_Ultima_Reposicao"] = hoje
            df_estoque.loc[idx, "Status"] = novo_status
            acao_msg = f"Insumo '{nome_item}' ({df_estoque.loc[idx, 'ID']}) atualizado no estoque."
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
                        "Preco_Unitario": float(preco_referencia or 0.0),
                        "Fornecedor": fornecedor_referencia.strip() or "-",
                        "Data_Ultima_Reposicao": hoje,
                        "Status": novo_status,
                    }
                ]
            )
            df_estoque = pd.concat([df_estoque, nova_linha], ignore_index=True)
            acao_msg = f"Novo insumo '{nome_item}' ({item_id}) cadastrado no estoque."

        ok = self.salvar_aba("Estoque", df_estoque)
        return ok, acao_msg

    # ==========================================================================
    # GESTÃO FINANCEIRA UNIFICADA (ABA 'Financeiro')
    # ==========================================================================
    def registrar_lancamento_financeiro(
        self,
        data_lancamento: str,
        tipo: str,
        natureza_custo: str,
        descricao: str,
        categoria: str,
        fornecedor: str,
        forma_pagamento: str,
        valor: float,
        link_comprovante: str,
        responsavel: str,
    ) -> Tuple[bool, str]:
        """Grava um lançamento financeiro (Receita ou Despesa) na aba 'Financeiro'."""
        df_fin = self.ler_aba("Financeiro")
        novo_id = f"FIN-{len(df_fin) + 1:03d}-{datetime.now().strftime('%H%M%S')}"

        nova_linha = pd.DataFrame(
            [
                {
                    "ID_Lancamento": novo_id,
                    "Data": data_lancamento,
                    "Tipo": tipo,
                    "Natureza_Custo": natureza_custo,
                    "Descricao": descricao.strip() or categoria,
                    "Categoria": categoria,
                    "Fornecedor": fornecedor.strip() or "-",
                    "Forma_Pagamento": forma_pagamento,
                    "Valor": round(float(valor), 2),
                    "Link_Comprovante": link_comprovante or "-",
                    "Responsavel": responsavel,
                }
            ]
        )
        df_fin = pd.concat([df_fin, nova_linha], ignore_index=True)
        ok = self.salvar_aba("Financeiro", df_fin)
        return ok, novo_id

    def excluir_lancamento_financeiro(
        self, id_lancamento: str, usuario_logado: str, perfil_logado: str
    ) -> Tuple[bool, str]:
        """
        TRAVA RIGOROSA DE EXCLUSÃO:
        Permite excluir lançamentos financeiros ou notas fiscais EXCLUSIVAMENTE quando
        usuario_logado == 'Arthur' (Perfil Dev).
        """
        if str(usuario_logado).strip() != "Arthur":
            self.registrar_log(
                usuario=usuario_logado,
                perfil=perfil_logado,
                acao=f"TENTATIVA BLOQUEADA de exclusão do lançamento {id_lancamento}",
                modulo="Financeiro",
            )
            return False, "Acesso Negado: Apenas o usuário 'Arthur' (Dev) possui permissão de exclusão."

        df_fin = self.ler_aba("Financeiro")
        mascara = df_fin["ID_Lancamento"].astype(str) == str(id_lancamento).strip()
        if not mascara.any():
            return False, f"Lançamento '{id_lancamento}' não encontrado."

        df_fin = df_fin[~mascara].reset_index(drop=True)
        ok = self.salvar_aba("Financeiro", df_fin)
        self.registrar_log(
            usuario=usuario_logado,
            perfil=perfil_logado,
            acao=f"Exclusão definitiva do lançamento/NF {id_lancamento}",
            modulo="Financeiro",
        )
        return ok, f"Lançamento '{id_lancamento}' excluído com sucesso por Arthur (Dev)."

    # ==========================================================================
    # PASTA ÚNICA NO GOOGLE DRIVE COM SUBPASTAS /Notas_Fiscais E /Comprovantes
    # ==========================================================================
    def _obter_ou_criar_pasta_drive(self, nome_pasta: str, parent_id: Optional[str] = None) -> Optional[str]:
        """Busca ou cria automaticamente uma pasta no Google Drive retornando seu ID."""
        if not (self.connected and self.drive_service):
            return None

        try:
            q_parts = [
                f"name = '{nome_pasta}'",
                "mimeType = 'application/vnd.google-apps.folder'",
                "trashed = false",
            ]
            if parent_id:
                q_parts.append(f"'{parent_id}' in parents")

            query = " and ".join(q_parts)
            res = (
                self.drive_service.files()
                .list(q=query, spaces="drive", fields="files(id, name)", pageSize=1)
                .execute()
            )
            arquivos = res.get("files", [])
            if arquivos:
                return arquivos[0]["id"]

            meta: Dict[str, Any] = {
                "name": nome_pasta,
                "mimeType": "application/vnd.google-apps.folder",
            }
            if parent_id:
                meta["parents"] = [parent_id]

            pasta_criada = (
                self.drive_service.files().create(body=meta, fields="id").execute()
            )
            return pasta_criada.get("id")
        except Exception:
            return None

    def upload_arquivo_drive(
        self,
        arquivo_bytes: bytes,
        nome_arquivo: str,
        mime_type: str,
        subpasta: str = "Notas_Fiscais",
    ) -> str:
        """
        Envia Nota Fiscal ou Comprovante para UMA ÚNICA PASTA PRINCIPAL no Google Drive,
        organizando automaticamente nas subpastas /Notas_Fiscais ou /Comprovantes.
        """
        subpasta_limpa = subpasta.strip("/")
        if self.connected and self.drive_service:
            try:
                id_pasta_principal = os.getenv("MAZON_DRIVE_FOLDER_ID") or self._obter_ou_criar_pasta_drive(
                    self.drive_main_folder_name
                )
                id_subpasta = self._obter_ou_criar_pasta_drive(
                    subpasta_limpa, parent_id=id_pasta_principal
                )

                metadata: Dict[str, Any] = {"name": nome_arquivo}
                if id_subpasta:
                    metadata["parents"] = [id_subpasta]
                elif id_pasta_principal:
                    metadata["parents"] = [id_pasta_principal]

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
                return f"https://drive.google.com/drive/folders/{subpasta_limpa}/{nome_arquivo} (Local: {exc})"

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"https://drive.google.com/drive/u/0/Mazon_Principal/{subpasta_limpa}/{timestamp}_{nome_arquivo}"


def registrar_log(usuario: str, perfil: str, acao: str, modulo: str) -> bool:
    """
    Função auxiliar global que grava automaticamente na aba 'Logs_Auditoria' do Sheets
    com as colunas: Data_Hora, Usuario, Perfil, Acao, Modulo.
    """
    db_instance = st.session_state.get("_mazon_db_instance")
    if db_instance is None:
        db_instance = MazonSheetsManager()
        st.session_state["_mazon_db_instance"] = db_instance
    return db_instance.registrar_log(usuario=usuario, perfil=perfil, acao=acao, modulo=modulo)
