"""
RESTAURANTE MAZON - Web App Unificado de Gestão de Estoque, Inteligência Financeira, RBAC e Auditoria
Banco de Dados Centralizado em Planilha Única no Google Sheets ("Mazon - Banco de Dados")
e Upload de Documentos em Pasta Única no Google Drive (/Notas_Fiscais e /Comprovantes).
"""

from __future__ import annotations

import io
import os
from datetime import date, datetime, timedelta
from typing import Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Image as RLImage,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from sheets_manager import (
    CATEGORIAS_FINANCEIRAS,
    CATEGORIAS_OFICIAIS,
    FORMAS_PAGAMENTO,
    MazonSheetsManager,
    registrar_log,
)


def resolver_caminho_recurso(nome_arquivo: str) -> str:
    """Retorna o caminho válido do arquivo no diretório atual ou no pacote."""
    if os.path.exists(nome_arquivo):
        return nome_arquivo
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), nome_arquivo)


# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA E FAVICON (logo2.png)
# ==============================================================================
_icone_aba = Image.open("logo2.png")

st.set_page_config(
    page_title="Restaurante Mazon | Gestão de Estoque & Inteligência Financeira",
    page_icon=_icone_aba,
    layout="wide",
    initial_sidebar_state="expanded",
)


def obter_caminho_logo() -> Optional[str]:
    """Busca dinamicamente a logo oficial entre os formatos suportados (.jpeg, .jpg, .png)."""
    formatos = [
        "logo.jpeg",
        "logo.jpg",
        "logo.png",
        "WhatsApp Image 2026-11-09 at 17.04.39.jpeg",
    ]
    for fmt in formatos:
        if os.path.exists(fmt):
            return fmt
        caminho_bundle = os.path.join(os.path.dirname(os.path.abspath(__file__)), fmt)
        if os.path.exists(caminho_bundle):
            return caminho_bundle
    return None


def injetar_estilo_luxury_minimalist() -> None:
    """
    Injeta a fonte Google Fonts Poppins restrita a elementos de texto,
    protege os ícones Material do Streamlit e aplica o style.css.
    """
    regras_fonte_e_icones = (
        "<style>"
        "@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');"
        "html, body, p, h1, h2, h3, h4, h5, h6, label, input, button, select, textarea, td, th, .stMarkdown p {"
        "font-family: 'Poppins', sans-serif !important;"
        "}"
        "[data-testid='stIcon'], [data-testid='stIconMaterial'], span[data-testid='stIconMaterial'], "
        "button [data-testid='stIconMaterial'], summary [data-testid='stIconMaterial'], "
        "header [data-testid='stIconMaterial'], i, .material-symbols-outlined, .material-icons, "
        "[class*='st-emotion-cache'] i {"
        "font-family: 'Material Symbols Outlined', 'Material Symbols Rounded', 'Material Icons' !important;"
        "font-weight: normal !important;"
        "font-style: normal !important;"
        "}"
        "</style>"
    )
    st.markdown(regras_fonte_e_icones, unsafe_allow_html=True)

    caminho_css = resolver_caminho_recurso("style.css")
    if os.path.exists(caminho_css):
        with open(caminho_css, "r", encoding="utf-8") as f:
            css_content = f.read()
        st.markdown(f"<style>{css_content}</style>", unsafe_allow_html=True)

    st.markdown(
        """<style>
/* Esconde totalmente a bolinha/círculo do st.radio (todas as versões do Streamlit) */
div[data-testid="stRadioButton"] [role="radio"],
div[data-testid="stRadioButton"] label > div:first-child,
div[data-baseweb="radio"] > div:first-child,
label[data-baseweb="radio"] > div:first-child {
    display: none !important;
    visibility: hidden !important;
    width: 0px !important;
    height: 0px !important;
}

/* Remaneja o alinhamento do texto */
div[data-testid="stRadioButton"] label,
label[data-baseweb="radio"] {
    padding-left: 5px !important;
    margin-left: 0px !important;
    cursor: pointer !important;
}

/* Destaca o item selecionado */
div[data-testid="stRadioButton"] label:has(input:checked),
label[data-baseweb="radio"]:has(input:checked) {
    font-weight: bold !important;
}
</style>""",
        unsafe_allow_html=True,
    )


injetar_estilo_luxury_minimalist()

# Inicializa o Gerenciador Unificado Google Sheets / Drive
if "_mazon_db_instance" not in st.session_state:
    st.session_state["_mazon_db_instance"] = MazonSheetsManager()
db: MazonSheetsManager = st.session_state["_mazon_db_instance"]


def formatar_moeda(valor: float) -> str:
    """Formata float no padrão monetário brasileiro R$ 1.234,56."""
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def renderizar_badge_status(status: str) -> str:
    """Retorna HTML do badge colorido segundo o status do estoque."""
    if status == "CRÍTICO/FALTA":
        return '<span class="badge-critico">CRÍTICO / FALTA</span>'
    elif status == "ALERTA":
        return '<span class="badge-alerta">ALERTA</span>'
    elif status == "PENDENTE":
        return '<span class="badge-pendente">PENDENTE</span>'
    return '<span class="badge-ok">OK / CHEIO</span>'


def gerar_pdf_lista_compras(df_compras: pd.DataFrame, custo_total_compra: float) -> bytes:
    """Gera o relatório em PDF (ReportLab) harmonizado na paleta Luxury Cream & Deep Navy."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title="Lista de Compras - Restaurante Mazon",
    )

    def _decorar_pagina_pdf(canvas_obj, doc_obj) -> None:
        canvas_obj.saveState()
        largura_pag, altura_pag = A4

        canvas_obj.setFillColor(colors.HexColor("#FAF6F0"))
        canvas_obj.rect(0, 0, largura_pag, altura_pag, stroke=0, fill=1)

        canvas_obj.setFillColor(colors.HexColor("#0B1B2B"))
        canvas_obj.rect(0, altura_pag - 6 * mm, largura_pag, 6 * mm, stroke=0, fill=1)
        canvas_obj.setStrokeColor(colors.HexColor("#D4A24C"))
        canvas_obj.setLineWidth(1.5)
        canvas_obj.line(0, altura_pag - 6 * mm, largura_pag, altura_pag - 6 * mm)

        canvas_obj.setFillColor(colors.HexColor("#687787"))
        canvas_obj.setFont("Helvetica", 8)
        canvas_obj.drawRightString(
            largura_pag - 14 * mm,
            9 * mm,
            f"Restaurante Mazon | Página {doc_obj.page}",
        )
        canvas_obj.restoreState()

    estilos = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "MazonTitulo",
        parent=estilos["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=19,
        alignment=1,
        textColor=colors.HexColor("#0B1B2B"),
        spaceAfter=4,
    )
    estilo_subtitulo = ParagraphStyle(
        "MazonSubtitulo",
        parent=estilos["Normal"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=13,
        alignment=1,
        textColor=colors.HexColor("#687787"),
        spaceAfter=12,
    )
    estilo_grupo = ParagraphStyle(
        "MazonGrupo",
        parent=estilos["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=15,
        textColor=colors.HexColor("#0B1B2B"),
        spaceBefore=10,
        spaceAfter=6,
    )

    elementos = []

    caminho_logo = obter_caminho_logo()
    if caminho_logo and os.path.exists(caminho_logo):
        try:
            img_logo = RLImage(caminho_logo, width=140, height=90)
            img_logo.hAlign = "CENTER"
            elementos.append(img_logo)
            elementos.append(Spacer(1, 4 * mm))
        except Exception:
            pass

    data_emissao = datetime.now().strftime("%d/%m/%Y às %H:%M")

    elementos.append(Paragraph("RESTAURANTE MAZON — LISTA OFICIAL DE COMPRAS", estilo_titulo))
    elementos.append(
        Paragraph(
            f"<b>Data de Emissão:</b> {data_emissao} &nbsp;|&nbsp; "
            f"<b>Total de Itens:</b> {len(df_compras)} &nbsp;|&nbsp; "
            f"<b>Orçamento Estimado:</b> {formatar_moeda(custo_total_compra)}",
            estilo_subtitulo,
        )
    )

    cabecalho_tabela = [
        "Insumo",
        "Categoria",
        "Qtd Atual",
        "Estoque Mínimo",
        "Estoque Máximo",
        "Qtd a Comprar",
    ]

    df_ordenado = df_compras.sort_values(["Categoria", "Item"])

    for categoria, grupo in df_ordenado.groupby("Categoria", sort=False):
        elementos.append(Paragraph(f"Categoria: {categoria} ({len(grupo)} itens)", estilo_grupo))

        dados_tabela = [cabecalho_tabela]
        for _, row in grupo.iterrows():
            unid = str(row["Unidade"])
            dados_tabela.append(
                [
                    str(row["Item"]),
                    str(row["Categoria"]),
                    f"{float(row['Qtd_Atual']):.2f} {unid}",
                    f"{float(row['Est_Minimo']):.2f} {unid}",
                    f"{float(row['Est_Maximo']):.2f} {unid}",
                    f"{float(row['Qtd_Necessaria']):.2f} {unid}",
                ]
            )

        tabela = Table(
            dados_tabela,
            colWidths=[48 * mm, 44 * mm, 22 * mm, 24 * mm, 24 * mm, 24 * mm],
            repeatRows=1,
        )
        tabela.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B1B2B")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#F4EDE3")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 8.5),
                    ("LINEBELOW", (0, 0), (-1, 0), 1.5, colors.HexColor("#D4A24C")),
                    ("ALIGN", (0, 0), (1, -1), "LEFT"),
                    ("ALIGN", (2, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 1), (-1, -1), 8.5),
                    ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor("#0B1B2B")),
                    ("FONTNAME", (5, 1), (5, -1), "Helvetica-Bold"),
                    ("TEXTCOLOR", (5, 1), (5, -1), colors.HexColor("#9A6B1F")),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F4EDE3")]),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#EBE4D8")),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        elementos.append(tabela)
        elementos.append(Spacer(1, 4 * mm))

    doc.build(
        elementos,
        onFirstPage=_decorar_pagina_pdf,
        onLaterPages=_decorar_pagina_pdf,
    )
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


# ==============================================================================
# ESTADO DE AUTENTICAÇÃO E TELA DE LOGIN INICIAL (BLOQUEIO TOTAL ANTES DO LOGIN)
# ==============================================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "usuario" not in st.session_state:
    st.session_state["usuario"] = None
if "perfil" not in st.session_state:
    st.session_state["perfil"] = None


def renderizar_tela_login() -> None:
    """Renderiza a tela central de autenticação bloqueando qualquer conteúdo antes do login."""
    _, col_centro, _ = st.columns([1.2, 1.4, 1.2])
    with col_centro:
        st.markdown("<div style='height: 36px;'></div>", unsafe_allow_html=True)
        caminho_logo = obter_caminho_logo()
        if caminho_logo:
            _, c_img, _ = st.columns([1, 2, 1])
            with c_img:
                st.image(caminho_logo, use_container_width=True)

        st.markdown(
            '<div class="mazon-login-box">'
            '<h2 style="margin:0 0 4px 0; color:#0B1B2B; font-size:1.45rem;">Acesso Restrito — Restaurante Mazon</h2>'
            '<p style="color:#687787; font-size:0.86rem; margin-bottom:0;">Insira suas credenciais para acessar o Sistema Unificado de Estoque &amp; Inteligência Financeira</p>'
            "</div>",
            unsafe_allow_html=True,
        )

        with st.form("form_login_mazon", clear_on_submit=False):
            in_usuario = st.text_input("Usuário", placeholder="Ex: Arthur, Claudio, Caixa, Equipe")
            in_senha = st.text_input("Senha", type="password", placeholder="Digite sua senha")
            btn_entrar = st.form_submit_button("Entrar no Sistema", use_container_width=True)

            if btn_entrar:
                dados_auth = db.autenticar_usuario(in_usuario, in_senha)
                if dados_auth:
                    st.session_state["autenticado"] = True
                    st.session_state["usuario"] = dados_auth["usuario"]
                    st.session_state["perfil"] = dados_auth["perfil"]
                    registrar_log(
                        usuario=dados_auth["usuario"],
                        perfil=dados_auth["perfil"],
                        acao=f"Login realizado com sucesso no sistema ({dados_auth['perfil']})",
                        modulo="Autenticacao",
                    )
                    st.rerun()
                else:
                    registrar_log(
                        usuario=in_usuario.strip() or "Desconhecido",
                        perfil="Nao_Autenticado",
                        acao=f"Tentativa falha de login para o usuário '{in_usuario}'",
                        modulo="Autenticacao",
                    )
                    st.error("Usuário ou senha inválidos. Verifique suas credenciais.")


if not st.session_state["autenticado"]:
    renderizar_tela_login()
    st.stop()


# ==============================================================================
# USUÁRIO AUTENTICADO: VARIÁVEIS DE SESSÃO E SIDEBAR
# ==============================================================================
USUARIO_ATUAL: str = str(st.session_state["usuario"])
PERFIL_ATUAL: str = str(st.session_state["perfil"])

caminho_logo_sb = obter_caminho_logo()
if caminho_logo_sb:
    st.sidebar.image(caminho_logo_sb, use_container_width=True)
else:
    st.sidebar.markdown(
        '<div style="text-align: center; padding: 10px 0 16px 0;">'
        '<div style="font-family: \'Poppins\', sans-serif; font-size: 1.5rem; font-weight: 600; color: #F4EDE3; letter-spacing: 2px;">MAZON</div>'
        '<div style="font-size: 0.72rem; color: #D4A24C; text-transform: uppercase; letter-spacing: 1.4px; margin-top: 2px;">Restaurante Japonês</div>'
        "</div>",
        unsafe_allow_html=True,
    )

with st.sidebar:
    # Identificação do Usuário Logado e Botão Sair (Logout)
    st.markdown(
        f'<div style="background: rgba(255,255,255,0.05); border: 1px solid rgba(212,162,76,0.35); border-radius: 10px; padding: 12px 14px; margin-bottom: 10px;">'
        f'<div style="font-size: 0.72rem; text-transform: uppercase; letter-spacing: 1px; color: #D4A24C;">Sessão Ativa · Perfil {PERFIL_ATUAL}</div>'
        f'<div style="font-size: 1.02rem; font-weight: 600; color: #F4EDE3; margin-top: 2px;">👤 {USUARIO_ATUAL}</div>'
        "</div>",
        unsafe_allow_html=True,
    )

    if st.button("🚪 Sair (Logout)", use_container_width=True):
        registrar_log(
            usuario=USUARIO_ATUAL,
            perfil=PERFIL_ATUAL,
            acao="Logout realizado pelo usuário",
            modulo="Autenticacao",
        )
        st.session_state["autenticado"] = False
        st.session_state["usuario"] = None
        st.session_state["perfil"] = None
        st.rerun()

    st.markdown('<hr style="border-color: rgba(255,255,255,0.08); margin: 14px 0;">', unsafe_allow_html=True)

    # Status da Conexão Google Cloud
    cor_status_nuvem = "#1F9D55" if db.connected else "#D4A24C"
    titulo_nuvem = "Mazon - Banco de Dados Online" if db.connected else "Banco Unificado Ativo"
    st.markdown(
        f'<div style="padding: 8px 12px; margin-bottom: 12px; border-left: 2px solid {cor_status_nuvem};">'
        f'<div style="font-size: 0.74rem; font-weight: 500; color: {cor_status_nuvem};">● {titulo_nuvem}</div>'
        f'<div style="font-size: 0.71rem; color: #C5D1DE; opacity: 0.78; margin-top: 2px;">{db.connection_message}</div>'
        "</div>",
        unsafe_allow_html=True,
    )

    # Ajuste Rápido de Contagem de Estoque na Sidebar
    st.markdown(
        '<p style="font-size:0.7rem; text-transform:uppercase; letter-spacing:1.4px; color:#D4A24C; font-weight:500; margin: 10px 0 6px 4px; opacity:0.9;">⚡ Movimento Rápido de Estoque</p>',
        unsafe_allow_html=True,
    )
    df_est_sidebar = db.ler_aba("Estoque")
    opcoes_itens = df_est_sidebar["Item"].tolist() if not df_est_sidebar.empty else []
    if opcoes_itens:
        item_rapido = st.selectbox("Insumo", options=opcoes_itens, key="sb_item_rapido")
        linha_atual = df_est_sidebar[df_est_sidebar["Item"] == item_rapido].iloc[0]
        nova_qtd_rapida = st.number_input(
            f"Saldo Atual ({linha_atual['Unidade']})",
            min_value=0.0,
            value=float(linha_atual["Qtd_Atual"]),
            step=0.5,
            key="sb_qtd_rapida",
        )
        if st.button("Sincronizar Saldo", use_container_width=True):
            qtd_anterior = float(linha_atual["Qtd_Atual"])
            ok, _ = db.cadastrar_ou_atualizar_item_estoque(
                item_id=str(linha_atual["ID"]),
                nome_item=str(linha_atual["Item"]),
                categoria=str(linha_atual["Categoria"]),
                unidade=str(linha_atual["Unidade"]),
                qtd_atual=nova_qtd_rapida,
                est_minimo=float(linha_atual["Est_Minimo"]),
                est_maximo=float(linha_atual["Est_Maximo"]),
                preco_referencia=float(linha_atual.get("Preco_Unitario", 0.0)),
                fornecedor_referencia=str(linha_atual.get("Fornecedor", "-")),
            )
            if ok:
                registrar_log(
                    usuario=USUARIO_ATUAL,
                    perfil=PERFIL_ATUAL,
                    acao=f"Atualização rápida de estoque: '{item_rapido}' de {qtd_anterior} para {nova_qtd_rapida} {linha_atual['Unidade']}",
                    modulo="Estoque",
                )
                st.success("Estoque atualizado!")
                st.rerun()

    # EXCLUSIVO PARA PERFIL DEV (Arthur): Gestão e Cadastro de Usuários na Sidebar
    if PERFIL_ATUAL == "Dev" and USUARIO_ATUAL == "Arthur":
        st.markdown('<hr style="border-color: rgba(255,255,255,0.08); margin: 14px 0;">', unsafe_allow_html=True)
        st.markdown(
            '<p style="font-size:0.7rem; text-transform:uppercase; letter-spacing:1.4px; color:#D4A24C; font-weight:600; margin: 6px 0 8px 4px;">🔐 Gestão de Usuários (Dev)</p>',
            unsafe_allow_html=True,
        )
        with st.form("form_gestao_usuarios_dev", clear_on_submit=True):
            novo_u_nome = st.text_input("Nome do Usuário", placeholder="Ex: Marina")
            novo_u_perfil = st.selectbox("Perfil de Acesso", options=["Dev", "Gerente", "Caixa", "Equipe"])
            novo_u_senha = st.text_input("Senha de Acesso", type="password")
            novo_u_ativo = st.selectbox("Status", options=["Sim", "Não"])
            if st.form_submit_button("Salvar Usuário", use_container_width=True):
                ok_u, msg_u = db.cadastrar_ou_atualizar_usuario(
                    novo_usuario=novo_u_nome,
                    perfil=novo_u_perfil,
                    senha=novo_u_senha,
                    ativo=novo_u_ativo,
                )
                if ok_u:
                    registrar_log(
                        usuario=USUARIO_ATUAL,
                        perfil=PERFIL_ATUAL,
                        acao=f"Cadastro/atualização de usuário '{novo_u_nome}' com perfil '{novo_u_perfil}' (Ativo={novo_u_ativo})",
                        modulo="Usuarios",
                    )
                    st.success(msg_u)
                    st.rerun()
                else:
                    st.warning(msg_u)


# ==============================================================================
# FUNÇÕES DE RENDERIZAÇÃO DE CADA TAB
# ==============================================================================
def renderizar_tab_home() -> None:
    """
    🏠 Tab Home:
    Apresenta resumo executivo customizado conforme o perfil logado:
    - Dev / Gerente: Cartões de métricas financeiras e de estoque do dia + gráficos executivos.
    - Caixa: Atalhos de rotina de fechamento e resumo dos lançamentos do dia.
    - Equipe: Alertas operacionais de itens com estoque baixo/crítico ou pendentes de contagem.
    """
    hoje_str = datetime.now().strftime("%Y-%m-%d")
    df_estoque = db.obter_estoque_enriquecido()
    df_criticos = df_estoque[df_estoque["Status"].isin(["CRÍTICO/FALTA", "PENDENTE", "ALERTA"])].copy()

    st.markdown(
        f'<div style="margin-bottom: 18px;">'
        f'<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">Olá, {USUARIO_ATUAL} ({PERFIL_ATUAL})</h1>'
        f'<p style="color:#687787; font-size:0.9rem; margin:0;">Resumo executivo e operacional do Restaurante Mazon — {datetime.now().strftime("%d/%m/%Y")}</p>'
        "</div>",
        unsafe_allow_html=True,
    )

    # 1. VISÃO DEV & GERENTE
    if PERFIL_ATUAL in ["Dev", "Gerente"]:
        df_fin = db.ler_aba("Financeiro").copy()
        df_fin["Valor"] = pd.to_numeric(df_fin["Valor"], errors="coerce").fillna(0.0)

        df_hoje = df_fin[df_fin["Data"].astype(str) == hoje_str]
        fat_hoje = float(df_hoje[df_hoje["Tipo"] == "Receita"]["Valor"].sum())
        desp_hoje = float(df_hoje[df_hoje["Tipo"] == "Despesa"]["Valor"].sum())
        fat_acumulado = float(df_fin[df_fin["Tipo"] == "Receita"]["Valor"].sum())
        desp_acumulada = float(df_fin[df_fin["Tipo"] == "Despesa"]["Valor"].sum())
        saldo_liquido = fat_acumulado - desp_acumulada

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Faturamento de Hoje</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(fat_hoje)}</div>'
                f'<div class="mazon-kpi-sub">Despesas hoje: {formatar_moeda(desp_hoje)}</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with c2:
            cor_saldo = "#1F9D55" if saldo_liquido >= 0 else "#D9383A"
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Resultado Líquido Acumulado</div>'
                f'<div class="mazon-kpi-value" style="color:{cor_saldo};">{formatar_moeda(saldo_liquido)}</div>'
                f'<div class="mazon-kpi-sub">Receitas: {formatar_moeda(fat_acumulado)}</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with c3:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Insumos em Alerta / Pendentes</div>'
                f'<div class="mazon-kpi-value" style="color:#D9383A;">{len(df_criticos)}</div>'
                f'<div class="mazon-kpi-sub">De {len(df_estoque)} insumos oficiais monitorados</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with c4:
            valor_est = float(df_estoque["Valor_Total_Item"].sum())
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Capital em Estoque</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(valor_est)}</div>'
                '<div class="mazon-kpi-sub">Avaliado pelo custo unitário cadastrado</div>'
                "</div>",
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)

        col_g1, col_g2 = st.columns([3, 2])
        with col_g1:
            st.markdown(
                '<div class="mazon-card" style="margin-bottom: 8px; padding: 1.15rem 1.4rem 0.4rem 1.4rem;">'
                '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">Evolução Recente: Receitas vs. Despesas (R$)</h4>'
                "</div>",
                unsafe_allow_html=True,
            )
            if not df_fin.empty:
                df_resumo_dia = (
                    df_fin.groupby(["Data", "Tipo"], as_index=False)["Valor"].sum()
                )
                fig_home = px.bar(
                    df_resumo_dia,
                    x="Data",
                    y="Valor",
                    color="Tipo",
                    barmode="group",
                    color_discrete_map={"Receita": "#0B1B2B", "Despesa": "#D4A24C"},
                )
                fig_home.update_layout(
                    paper_bgcolor="#FFFFFF",
                    plot_bgcolor="#FFFFFF",
                    font=dict(color="#0B1B2B", family="Poppins"),
                    xaxis=dict(title="Data", gridcolor="#EBE4D8"),
                    yaxis=dict(title="Valor (R$)", gridcolor="#EBE4D8"),
                    margin=dict(l=12, r=12, t=15, b=15),
                    height=330,
                )
                st.plotly_chart(fig_home, use_container_width=True)
        with col_g2:
            st.markdown(
                '<div class="mazon-card" style="margin-bottom: 8px; padding: 1.15rem 1.4rem 0.4rem 1.4rem;">'
                '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">Distribuição de Insumos por Categoria</h4>'
                "</div>",
                unsafe_allow_html=True,
            )
            df_contagem_cat = (
                df_estoque.groupby("Categoria", as_index=False)["ID"]
                .count()
                .rename(columns={"ID": "Qtd_Itens"})
            )
            fig_pie = px.pie(
                df_contagem_cat,
                names="Categoria",
                values="Qtd_Itens",
                hole=0.6,
                color_discrete_sequence=[
                    "#0B1B2B",
                    "#D4A24C",
                    "#1C354D",
                    "#C59237",
                    "#34506B",
                    "#E0B86B",
                    "#526375",
                    "#8A6D3B",
                    "#B8ADA0",
                ],
            )
            fig_pie.update_traces(textinfo="percent+label", textfont_color="#0B1B2B")
            fig_pie.update_layout(
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
                font=dict(color="#0B1B2B", family="Poppins"),
                showlegend=False,
                margin=dict(l=10, r=10, t=10, b=10),
                height=330,
            )
            st.plotly_chart(fig_pie, use_container_width=True)

    # 2. VISÃO CAIXA (ATALHOS DE ROTINA OPERACIONAL)
    elif PERFIL_ATUAL == "Caixa":
        df_fin_caixa = db.ler_aba("Financeiro")
        lancamentos_hoje_caixa = df_fin_caixa[
            (df_fin_caixa["Data"].astype(str) == hoje_str)
            & (df_fin_caixa["Responsavel"].astype(str) == USUARIO_ATUAL)
        ]

        c_cx1, c_cx2, c_cx3 = st.columns(3)
        with c_cx1:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Seus Lançamentos Hoje</div>'
                f'<div class="mazon-kpi-value">{len(lancamentos_hoje_caixa)}</div>'
                '<div class="mazon-kpi-sub">Registros enviados para a planilha central</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with c_cx2:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Rotina 1 · Fechamento Diário</div>'
                '<div class="mazon-kpi-value" style="font-size:1.25rem;">💰 Lançamentos / Caixa</div>'
                '<div class="mazon-kpi-sub">Acesse a aba superior para informar o Faturamento Bruto</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with c_cx3:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Rotina 2 · Notas &amp; Comprovantes</div>'
                '<div class="mazon-kpi-value" style="font-size:1.25rem;">📎 Upload Google Drive</div>'
                '<div class="mazon-kpi-sub">Anexe PDFs e fotos de NFs na aba Lançamentos / Caixa</div>'
                "</div>",
                unsafe_allow_html=True,
            )

        st.markdown(
            '<div class="mazon-card" style="margin-top:14px;">'
            '<h4 style="margin:0 0 8px 0; color:#0B1B2B;">📋 Checklist Operacional do Turno (Perfil Caixa)</h4>'
            '<ul style="color:#526375; font-size:0.88rem; line-height:1.7; margin-bottom:0;">'
            "<li><strong>1. Lançamento do Faturamento Bruto Diário:</strong> Acesse a aba <code>💰 Lançamentos / Caixa</code> e registre o fechamento de Salão e Delivery por forma de pagamento.</li>"
            "<li><strong>2. Digitalização de Notas Fiscais e Comprovantes:</strong> Anexe os comprovantes e notas de fornecedores recebidos no turno; o sistema arquiva automaticamente nas pastas <code>/Notas_Fiscais</code> e <code>/Comprovantes</code> do Google Drive.</li>"
            "<li><strong>3. Conferência de Insumos:</strong> Utilize a aba <code>📦 Estoque</code> para consultar ou informar entradas de mercadorias recebidas.</li>"
            "</ul>"
            "</div>",
            unsafe_allow_html=True,
        )

    # 3. ALERTAS DE ESTOQUE BAIXO / PENDENTE (EXIBIDO PARA EQUIPE E TODOS OS PERFIS)
    if not df_criticos.empty:
        pills_list = []
        for _, row in df_criticos.iterrows():
            deficit = max(0.0, float(row["Est_Minimo"]) - float(row["Qtd_Atual"]))
            badge_card = renderizar_badge_status(str(row["Status"]))
            pills_list.append(
                '<div class="critical-item-pill">'
                '<div style="display:flex; justify-content:space-between; align-items:center;">'
                f'<span style="font-weight:600; color:#0B1B2B; font-size:0.88rem;">{row["Item"]}</span>'
                f"{badge_card}"
                "</div>"
                f'<div style="font-size:0.74rem; color:#9A6B1F; margin-top:3px;">{row["Categoria"]} · {row["ID"]}</div>'
                '<div style="font-size:0.8rem; color:#0B1B2B; margin-top:5px;">'
                f'Atual: <strong style="color:#D9383A;">{float(row["Qtd_Atual"]):.1f} {row["Unidade"]}</strong>'
                f' &nbsp;/&nbsp; Mín: <strong>{float(row["Est_Minimo"]):.1f} {row["Unidade"]}</strong>'
                f" &nbsp;(Déficit: -{deficit:.1f} {row['Unidade']})"
                "</div>"
                "</div>"
            )
        pills_html = "".join(pills_list)

        st.markdown(
            '<div class="critical-hero-card">'
            '<div class="critical-hero-header">'
            f'<div class="critical-hero-title">⚠️ Alertas de Estoque Baixo ou Pendentes de Contagem ({len(df_criticos)} itens)</div>'
            '<span class="badge-best-price">Atenção Operacional</span>'
            "</div>"
            '<p style="color:#687787; font-size:0.85rem; margin:0;">'
            "Verifique os insumos abaixo para reposição imediata ou atualização de contagem física:"
            "</p>"
            f'<div class="critical-items-grid">{pills_html}</div>'
            "</div>",
            unsafe_allow_html=True,
        )


def renderizar_tab_estoque() -> None:
    """
    📦 Tab Estoque:
    - Mantém a tabela de insumos e movimentações de estoque (entradas/saídas) integrada à planilha.
    - Permite edição de cadastro (novo item / parâmetros de mínimo, máximo e preço) APENAS para Dev e Gerente.
    - Inclui o Gerador de Lista de Compras (PDF ReportLab + CSV).
    """
    pode_editar_cadastro = PERFIL_ATUAL in ["Dev", "Gerente"]

    st.markdown(
        '<div style="margin-bottom: 16px;">'
        '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">📦 Gestão de Estoque &amp; Movimentações</h1>'
        '<p style="color:#687787; font-size:0.9rem; margin:0;">Controle dos 90 insumos oficiais sincronizados com a aba <code>Estoque</code> da planilha unificada.</p>'
        "</div>",
        unsafe_allow_html=True,
    )

    df_estoque = db.obter_estoque_enriquecido()

    # Barra de filtros e botões de ação conforme permissão RBAC
    col_f1, col_f2, col_f3, col_btn = st.columns([2, 2, 1.8, 1.6])
    with col_f1:
        busca_texto = st.text_input("Buscar por Item ou ID", placeholder="Ex: Salmão, H01, MZ-001...")
    with col_f2:
        filtro_cat = st.multiselect(
            "Filtrar por Categoria",
            options=CATEGORIAS_OFICIAIS,
            default=[],
            placeholder="Todas as categorias",
        )
    with col_f3:
        filtro_status = st.multiselect(
            "Filtrar por Status",
            options=["PENDENTE", "CRÍTICO/FALTA", "ALERTA", "OK/CHEIO"],
            default=[],
            placeholder="Todos os status",
        )
    with col_btn:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        if pode_editar_cadastro:
            abrir_modal_cadastro = st.button("＋ Cadastrar / Editar Item", use_container_width=True)
        else:
            abrir_modal_cadastro = False
            st.markdown(
                '<div style="font-size:0.76rem; color:#687787; padding-top:8px;">🔒 Edição de cadastro restrita a Dev/Gerente</div>',
                unsafe_allow_html=True,
            )

    # Modal de Cadastro / Edição de Parâmetros (Exclusivo para Dev e Gerente)
    @st.dialog("Cadastrar ou Editar Cadastro de Insumo")
    def modal_cadastro_estoque() -> None:
        st.markdown(
            "Preencha os parâmetros do insumo. Se o **ID** ou **Nome do Item** já existir, o cadastro será atualizado."
        )
        with st.form("form_modal_estoque_rbac", clear_on_submit=False):
            c1, c2 = st.columns(2)
            with c1:
                in_id = st.text_input("ID do Insumo (opcional para novo)", placeholder="Ex: MZ-001")
                in_nome = st.text_input("Nome do Item *", placeholder="Ex: Salmão")
                in_cat = st.selectbox("Categoria *", options=CATEGORIAS_OFICIAIS)
                in_unid = st.selectbox(
                    "Unidade de Medida *",
                    options=["kg", "un", "L", "pct", "maço", "par", "cx", "cartela", "vidro", "rolo", "fardo"],
                )
            with c2:
                in_qtd = st.number_input("Quantidade Atual *", min_value=0.0, value=0.0, step=0.5)
                in_min = st.number_input("Estoque Mínimo *", min_value=0.0, value=0.0, step=0.5)
                in_max = st.number_input("Estoque Máximo *", min_value=0.0, value=0.0, step=1.0)
                in_preco = st.number_input("Preço Unitário (R$)", min_value=0.0, value=0.0, step=1.0)
            in_forn = st.text_input("Fornecedor Principal", value="-")

            if st.form_submit_button("Salvar Cadastro na Planilha Unificada", use_container_width=True):
                if not in_nome.strip():
                    st.error("Informe o nome do insumo.")
                else:
                    ok, msg = db.cadastrar_ou_atualizar_item_estoque(
                        item_id=in_id,
                        nome_item=in_nome,
                        categoria=in_cat,
                        unidade=in_unid,
                        qtd_atual=in_qtd,
                        est_minimo=in_min,
                        est_maximo=in_max,
                        preco_referencia=in_preco,
                        fornecedor_referencia=in_forn,
                    )
                    if ok:
                        registrar_log(
                            usuario=USUARIO_ATUAL,
                            perfil=PERFIL_ATUAL,
                            acao=f"Edição de cadastro no estoque: '{in_nome}' (Qtd={in_qtd}, Min={in_min}, Max={in_max}, R$={in_preco})",
                            modulo="Estoque",
                        )
                        st.success(msg)
                        st.rerun()

    if abrir_modal_cadastro and pode_editar_cadastro:
        modal_cadastro_estoque()

    # Formulário de Movimentação de Estoque (Entrada / Saída Operacional)
    with st.expander("🔄 Registrar Entrada ou Saída de Mercadoria (Movimentação de Estoque)", expanded=False):
        with st.form("form_movimentacao_estoque"):
            cm1, cm2, cm3, cm4 = st.columns([2.2, 1.4, 1.4, 2])
            with cm1:
                mov_item = st.selectbox("Insumo", options=df_estoque["Item"].tolist())
            with cm2:
                mov_tipo = st.selectbox("Tipo de Movimento", options=["Entrada (+)", "Saída / Baixa (-)"])
            with cm3:
                mov_qtd = st.number_input("Quantidade Movimentada", min_value=0.01, value=1.0, step=0.5)
            with cm4:
                mov_obs = st.text_input("Motivo / NF / Observação", placeholder="Ex: Recebimento NF 402 ou Consumo Cozinha")

            if st.form_submit_button("Confirmar Movimentação de Estoque"):
                linha_item = df_estoque[df_estoque["Item"] == mov_item].iloc[0]
                qtd_antiga = float(linha_item["Qtd_Atual"])
                if mov_tipo.startswith("Entrada"):
                    nova_qtd = qtd_antiga + float(mov_qtd)
                else:
                    nova_qtd = max(0.0, qtd_antiga - float(mov_qtd))

                ok_mov, _ = db.cadastrar_ou_atualizar_item_estoque(
                    item_id=str(linha_item["ID"]),
                    nome_item=str(linha_item["Item"]),
                    categoria=str(linha_item["Categoria"]),
                    unidade=str(linha_item["Unidade"]),
                    qtd_atual=nova_qtd,
                    est_minimo=float(linha_item["Est_Minimo"]),
                    est_maximo=float(linha_item["Est_Maximo"]),
                    preco_referencia=float(linha_item.get("Preco_Unitario", 0.0)),
                    fornecedor_referencia=str(linha_item.get("Fornecedor", "-")),
                )
                if ok_mov:
                    registrar_log(
                        usuario=USUARIO_ATUAL,
                        perfil=PERFIL_ATUAL,
                        acao=f"{mov_tipo} de {mov_qtd} {linha_item['Unidade']} em '{mov_item}' (Saldo: {qtd_antiga} -> {nova_qtd}) | Obs: {mov_obs or '-'}",
                        modulo="Estoque",
                    )
                    st.success(
                        f"Movimentação registrada! Novo saldo de '{mov_item}': {nova_qtd:.2f} {linha_item['Unidade']}."
                    )
                    st.rerun()

    # Aplicação dos filtros na tabela de estoque
    df_filtrado = df_estoque.copy()
    if busca_texto.strip():
        termo = busca_texto.strip().lower()
        df_filtrado = df_filtrado[
            df_filtrado["Item"].astype(str).str.lower().str.contains(termo)
            | df_filtrado["ID"].astype(str).str.lower().str.contains(termo)
        ]
    if filtro_cat:
        df_filtrado = df_filtrado[df_filtrado["Categoria"].isin(filtro_cat)]
    if filtro_status:
        df_filtrado = df_filtrado[df_filtrado["Status"].isin(filtro_status)]

    linhas_lista = []
    for _, row in df_filtrado.iterrows():
        badge_html = renderizar_badge_status(str(row["Status"]))
        linhas_lista.append(
            "<tr>"
            f'<td style="font-weight:600; color:#9A6B1F;">{row["ID"]}</td>'
            f'<td style="font-weight:500; color:#0B1B2B;">{row["Item"]}</td>'
            f'<td>{row["Categoria"]}</td>'
            f'<td>{row["Unidade"]}</td>'
            f'<td style="font-weight:600;">{float(row["Qtd_Atual"]):.2f}</td>'
            f'<td>{float(row["Est_Minimo"]):.2f}</td>'
            f'<td>{float(row["Est_Maximo"]):.2f}</td>'
            f'<td>{formatar_moeda(float(row.get("Preco_Unitario", 0.0)))}</td>'
            f"<td>{badge_html}</td>"
            f'<td style="color:#687787;">{row["Data_Ultima_Reposicao"]}</td>'
            "</tr>"
        )

    corpo_tabela_html = "".join(linhas_lista) if linhas_lista else (
        '<tr><td colspan="10" style="text-align:center; color:#687787; padding:18px;">Nenhum insumo encontrado para os filtros selecionados.</td></tr>'
    )

    tabela_estoque_html = (
        '<div class="mazon-table-wrapper">'
        '<table class="mazon-table">'
        "<thead>"
        "<tr>"
        "<th>ID</th>"
        "<th>Item</th>"
        "<th>Categoria</th>"
        "<th>Unidade</th>"
        "<th>Qtd Atual</th>"
        "<th>Mínimo</th>"
        "<th>Máximo</th>"
        "<th>Custo Unit.</th>"
        "<th>Status</th>"
        "<th>Última Reposição</th>"
        "</tr>"
        "</thead>"
        f"<tbody>{corpo_tabela_html}</tbody>"
        "</table>"
        "</div>"
    )
    st.markdown(tabela_estoque_html, unsafe_allow_html=True)

    # Exportação da Lista de Compras (PDF / CSV) integrada à aba Estoque
    with st.expander("🛒 Gerador de Lista de Compras & Exportação em PDF / CSV", expanded=False):
        df_alerta = df_estoque[
            df_estoque["Status"].isin(["CRÍTICO/FALTA", "ALERTA", "PENDENTE"])
        ].copy()
        df_alerta["Qtd_Necessaria"] = (
            df_alerta["Est_Maximo"] - df_alerta["Qtd_Atual"]
        ).clip(lower=0.0)
        df_alerta["Custo_Estimado_Reposicao"] = (
            df_alerta["Qtd_Necessaria"] * df_alerta["Preco_Unitario"]
        )
        custo_total_compra = float(df_alerta["Custo_Estimado_Reposicao"].sum())

        pdf_bytes = gerar_pdf_lista_compras(df_alerta, custo_total_compra)
        csv_bytes = df_alerta.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")

        cdl1, cdl2 = st.columns(2)
        with cdl1:
            st.download_button(
                label="📄 Exportar Lista de Compras (PDF)",
                data=pdf_bytes,
                file_name=f"lista_compras_mazon_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        with cdl2:
            st.download_button(
                label="📥 Exportar Lista de Compras (Planilha CSV / Excel)",
                data=csv_bytes,
                file_name=f"lista_compras_mazon_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
            )


def renderizar_tab_financeiro() -> None:
    """
    💰 Tab Gestão Financeira (Dev e Gerente) ou 💰 Lançamentos / Caixa (Caixa):
    - Visão Restrita (Perfil Caixa): Exibe APENAS o formulário de lançamento do faturamento bruto diário
      e upload de arquivos (NFs/Comprovantes) para a pasta do Google Drive. Omite gráficos, DRE e margens.
    - Visão Completa (Dev e Gerente): Painel de filtros interativos no topo (Período, Fornecedor, Faixa de Valor,
      Categoria/Forma de Pagamento), cartões de métricas consolidadas, gráficos Plotly (Fluxo de Caixa, DRE,
      Faturamento por Período, Despesas por Categoria) e tabela detalhada.
    - TRAVA RIGOROSA DE EXCLUSÃO: Botões para excluir lançamentos ou remover NFs são renderizados e executados
      EXCLUSIVAMENTE quando st.session_state['usuario'] == 'Arthur' (Perfil Dev).
    """
    # ==========================================================================
    # VISÃO RESTRITA: PERFIL CAIXA
    # ==========================================================================
    if PERFIL_ATUAL == "Caixa":
        st.markdown(
            '<div style="margin-bottom: 18px;">'
            '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">💰 Lançamentos de Caixa &amp; Envio de Comprovantes</h1>'
            '<p style="color:#687787; font-size:0.9rem; margin:0;">Registre o faturamento bruto diário e envie Notas Fiscais ou Comprovantes diretamente para a pasta oficial do Google Drive.</p>'
            "</div>",
            unsafe_allow_html=True,
        )

        col_f_caixa, col_up_caixa = st.columns([1.2, 1.0])

        with col_f_caixa:
            st.markdown('<div class="mazon-card">', unsafe_allow_html=True)
            st.markdown("#### 💵 Lançamento de Faturamento Bruto Diário")
            with st.form("form_caixa_faturamento", clear_on_submit=True):
                cx_data = st.date_input("Data do Faturamento", value=date.today())
                cx_cat = st.selectbox(
                    "Origem do Faturamento",
                    options=[
                        "Faturamento Salão",
                        "Faturamento Delivery (iFood / Próprio)",
                        "Eventos & Reservas",
                    ],
                )
                cx_forma = st.selectbox("Forma de Pagamento Principal", options=FORMAS_PAGAMENTO)
                cx_valor = st.number_input(
                    "Valor Bruto Apurado (R$) *", min_value=0.01, value=100.0, step=50.0
                )
                cx_desc = st.text_input(
                    "Observação / Turno", placeholder="Ex: Fechamento Turno Almoço / Jantar"
                )
                cx_arq = st.file_uploader(
                    "Anexar Comprovante de Fechamento (opcional - PDF/JPG/PNG)",
                    type=["pdf", "png", "jpg", "jpeg"],
                    key="up_caixa_fechamento",
                )

                if st.form_submit_button("Registrar Faturamento no Sistema", use_container_width=True):
                    link_comp = "-"
                    if cx_arq is not None:
                        link_comp = db.upload_arquivo_drive(
                            arquivo_bytes=cx_arq.read(),
                            nome_arquivo=cx_arq.name,
                            mime_type=cx_arq.type or "application/octet-stream",
                            subpasta="Comprovantes",
                        )
                        registrar_log(
                            usuario=USUARIO_ATUAL,
                            perfil=PERFIL_ATUAL,
                            acao=f"Upload de comprovante '{cx_arq.name}' para /Comprovantes no Google Drive",
                            modulo="Caixa",
                        )

                    ok_fin, novo_id = db.registrar_lancamento_financeiro(
                        data_lancamento=cx_data.strftime("%Y-%m-%d"),
                        tipo="Receita",
                        natureza_custo="Faturamento Bruto",
                        descricao=cx_desc or f"Fechamento {cx_cat}",
                        categoria=cx_cat,
                        fornecedor="Cliente Final",
                        forma_pagamento=cx_forma,
                        valor=cx_valor,
                        link_comprovante=link_comp,
                        responsavel=USUARIO_ATUAL,
                    )
                    if ok_fin:
                        registrar_log(
                            usuario=USUARIO_ATUAL,
                            perfil=PERFIL_ATUAL,
                            acao=f"Lançamento de Faturamento Bruto ({novo_id}): {formatar_moeda(cx_valor)} via {cx_forma}",
                            modulo="Caixa",
                        )
                        st.success(f"Faturamento {novo_id} registrado com sucesso!")
                        st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        with col_up_caixa:
            st.markdown('<div class="mazon-card">', unsafe_allow_html=True)
            st.markdown("#### 📎 Upload de Notas Fiscais e Comprovantes (Google Drive)")
            with st.form("form_caixa_upload_nf", clear_on_submit=True):
                nf_data = st.date_input("Data do Documento", value=date.today(), key="dt_nf_cx")
                nf_subpasta = st.selectbox(
                    "Pasta de Destino no Google Drive",
                    options=["Notas_Fiscais", "Comprovantes"],
                )
                nf_forn = st.text_input("Fornecedor / Emissor *", placeholder="Ex: Mar & Cia Imports")
                nf_cat = st.selectbox(
                    "Categoria da Despesa / Nota",
                    options=[c for c in CATEGORIAS_FINANCEIRAS if not c.startswith("Faturamento")],
                )
                nf_forma = st.selectbox("Forma de Pagamento", options=FORMAS_PAGAMENTO, key="fp_nf_cx")
                nf_valor = st.number_input("Valor Total do Documento (R$) *", min_value=0.01, value=50.0, step=10.0)
                nf_arquivo = st.file_uploader(
                    "Selecione o arquivo da Nota Fiscal ou Comprovante (.pdf, .jpg, .png) *",
                    type=["pdf", "png", "jpg", "jpeg"],
                    key="up_caixa_nf",
                )

                if st.form_submit_button("Enviar para o Google Drive & Registrar", use_container_width=True):
                    if nf_arquivo is None:
                        st.error("Selecione um arquivo PDF ou imagem para fazer o upload.")
                    elif not nf_forn.strip():
                        st.error("Informe o nome do fornecedor.")
                    else:
                        link_drive = db.upload_arquivo_drive(
                            arquivo_bytes=nf_arquivo.read(),
                            nome_arquivo=nf_arquivo.name,
                            mime_type=nf_arquivo.type or "application/octet-stream",
                            subpasta=nf_subpasta,
                        )
                        ok_nf, id_nf = db.registrar_lancamento_financeiro(
                            data_lancamento=nf_data.strftime("%Y-%m-%d"),
                            tipo="Despesa",
                            natureza_custo="Custo Variável",
                            descricao=f"Documento: {nf_arquivo.name}",
                            categoria=nf_cat,
                            fornecedor=nf_forn,
                            forma_pagamento=nf_forma,
                            valor=nf_valor,
                            link_comprovante=link_drive,
                            responsavel=USUARIO_ATUAL,
                        )
                        if ok_nf:
                            registrar_log(
                                usuario=USUARIO_ATUAL,
                                perfil=PERFIL_ATUAL,
                                acao=f"Upload de '{nf_arquivo.name}' em /{nf_subpasta} e registro de nota ({id_nf}) de '{nf_forn}' ({formatar_moeda(nf_valor)})",
                                modulo="Caixa",
                            )
                            st.success(f"Arquivo enviado para /{nf_subpasta} e registrado ({id_nf})!")
                            st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
        return

    # ==========================================================================
    # VISÃO COMPLETA: PERFIS DEV E GERENTE
    # ==========================================================================
    st.markdown(
        '<div style="margin-bottom: 16px;">'
        '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">💰 Gestão Financeira, DRE &amp; Controle de Notas Fiscais</h1>'
        '<p style="color:#687787; font-size:0.9rem; margin:0;">Análise interativa de fluxo de caixa, faturamento, custos fixos/variáveis e comprovantes arquivados no Google Drive.</p>'
        "</div>",
        unsafe_allow_html=True,
    )

    df_fin = db.ler_aba("Financeiro").copy()
    df_fin["Valor"] = pd.to_numeric(df_fin["Valor"], errors="coerce").fillna(0.0)
    df_fin["Data_dt"] = pd.to_datetime(df_fin["Data"], errors="coerce").dt.date

    # 1. PAINEL DE FILTROS INTERATIVOS NO TOPO (VISÃO DEV E GERENTE)
    st.markdown('<div class="mazon-card" style="padding: 1.15rem 1.4rem;">', unsafe_allow_html=True)
    st.markdown(
        '<div style="font-size:0.78rem; text-transform:uppercase; letter-spacing:1px; color:#9A6B1F; font-weight:600; margin-bottom:10px;">🔎 Painel de Filtros Financeiros Interativos</div>',
        unsafe_allow_html=True,
    )

    datas_validas = df_fin["Data_dt"].dropna()
    dt_min_padrao = datas_validas.min() if not datas_validas.empty else (date.today() - timedelta(days=30))
    dt_max_padrao = datas_validas.max() if not datas_validas.empty else date.today()
    if dt_min_padrao > dt_max_padrao:
        dt_min_padrao = dt_max_padrao

    fornecedores_opcoes = sorted(df_fin["Fornecedor"].astype(str).unique().tolist()) if not df_fin.empty else []
    categorias_opcoes = sorted(df_fin["Categoria"].astype(str).unique().tolist()) if not df_fin.empty else CATEGORIAS_FINANCEIRAS
    formas_opcoes = sorted(df_fin["Forma_Pagamento"].astype(str).unique().tolist()) if not df_fin.empty else FORMAS_PAGAMENTO

    val_min_global = float(df_fin["Valor"].min()) if not df_fin.empty else 0.0
    val_max_global = float(df_fin["Valor"].max()) if not df_fin.empty else 10000.0
    if val_max_global <= val_min_global:
        val_max_global = val_min_global + 1000.0

    f_col1, f_col2, f_col3, f_col4 = st.columns([1.4, 1.4, 1.5, 1.5])
    with f_col1:
        periodo_sel = st.date_input(
            "Período (Inicial — Final)",
            value=(dt_min_padrao, dt_max_padrao),
        )
    with f_col2:
        fornecedores_sel = st.multiselect(
            "Fornecedor / Origem",
            options=fornecedores_opcoes,
            default=[],
            placeholder="Todos os fornecedores",
        )
    with f_col3:
        categorias_sel = st.multiselect(
            "Categoria Financeira",
            options=categorias_opcoes,
            default=[],
            placeholder="Todas as categorias",
        )
    with f_col4:
        formas_sel = st.multiselect(
            "Forma de Pagamento",
            options=formas_opcoes,
            default=[],
            placeholder="Todas as formas",
        )

    faixa_valor = st.slider(
        "Faixa de Valor do Lançamento (R$)",
        min_value=0.0,
        max_value=float(max(val_max_global, 100.0)),
        value=(0.0, float(max(val_max_global, 100.0))),
        step=10.0,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # Aplicação dos filtros sobre df_fin
    df_filtrado = df_fin.copy()
    if isinstance(periodo_sel, (list, tuple)) and len(periodo_sel) == 2:
        dt_ini, dt_fim = periodo_sel[0], periodo_sel[1]
        df_filtrado = df_filtrado[
            (df_filtrado["Data_dt"] >= dt_ini) & (df_filtrado["Data_dt"] <= dt_fim)
        ]
    elif isinstance(periodo_sel, date):
        df_filtrado = df_filtrado[df_filtrado["Data_dt"] >= periodo_sel]

    if fornecedores_sel:
        df_filtrado = df_filtrado[df_filtrado["Fornecedor"].isin(fornecedores_sel)]
    if categorias_sel:
        df_filtrado = df_filtrado[df_filtrado["Categoria"].isin(categorias_sel)]
    if formas_sel:
        df_filtrado = df_filtrado[df_filtrado["Forma_Pagamento"].isin(formas_sel)]

    df_filtrado = df_filtrado[
        (df_filtrado["Valor"] >= faixa_valor[0]) & (df_filtrado["Valor"] <= faixa_valor[1])
    ]

    # 2. CARTÕES DE MÉTRICAS CONSOLIDADAS (REATIVOS AOS FILTROS)
    faturamento_bruto = float(df_filtrado[df_filtrado["Tipo"] == "Receita"]["Valor"].sum())
    df_despesas = df_filtrado[df_filtrado["Tipo"] == "Despesa"]
    custos_variaveis = float(
        df_despesas[df_despesas["Natureza_Custo"] == "Custo Variável"]["Valor"].sum()
    )
    custos_fixos = float(
        df_despesas[df_despesas["Natureza_Custo"] == "Custo Fixo"]["Valor"].sum()
    )
    despesas_totais = float(df_despesas["Valor"].sum())
    lucro_liquido = faturamento_bruto - despesas_totais
    margem_liquida = (lucro_liquido / faturamento_bruto * 100.0) if faturamento_bruto > 0 else 0.0

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(
            '<div class="mazon-kpi-card">'
            '<div class="mazon-kpi-label">Faturamento Bruto</div>'
            f'<div class="mazon-kpi-value">{formatar_moeda(faturamento_bruto)}</div>'
            f'<div class="mazon-kpi-sub">{len(df_filtrado[df_filtrado["Tipo"] == "Receita"])} lançamentos de receita</div>'
            "</div>",
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            '<div class="mazon-kpi-card">'
            '<div class="mazon-kpi-label">Custos Fixos &amp; Variáveis</div>'
            f'<div class="mazon-kpi-value" style="color:#D9383A;">{formatar_moeda(despesas_totais)}</div>'
            f'<div class="mazon-kpi-sub">Var: {formatar_moeda(custos_variaveis)} · Fixo: {formatar_moeda(custos_fixos)}</div>'
            "</div>",
            unsafe_allow_html=True,
        )
    with k3:
        cor_lucro = "#1F9D55" if lucro_liquido >= 0 else "#D9383A"
        st.markdown(
            '<div class="mazon-kpi-card">'
            '<div class="mazon-kpi-label">Lucro Líquido</div>'
            f'<div class="mazon-kpi-value" style="color:{cor_lucro};">{formatar_moeda(lucro_liquido)}</div>'
            '<div class="mazon-kpi-sub">Resultado operacional no período filtrado</div>'
            "</div>",
            unsafe_allow_html=True,
        )
    with k4:
        cor_margem = "#1F9D55" if margem_liquida >= 0 else "#D9383A"
        st.markdown(
            '<div class="mazon-kpi-card">'
            '<div class="mazon-kpi-label">Margem Líquida (%)</div>'
            f'<div class="mazon-kpi-value" style="color:{cor_margem};">{margem_liquida:.1f}%</div>'
            '<div class="mazon-kpi-sub">Lucro Líquido / Faturamento Bruto</div>'
            "</div>",
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)

    # 3. GRÁFICOS INTERATIVOS PLOTLY (Fluxo de Caixa, DRE Simplificada, Faturamento por Período, Despesas por Categoria)
    g_col1, g_col2 = st.columns(2)

    with g_col1:
        st.markdown(
            '<div class="mazon-card" style="margin-bottom:8px; padding:1.1rem 1.4rem 0.4rem 1.4rem;">'
            '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">📊 DRE Simplificada (Cascata de Resultado)</h4>'
            "</div>",
            unsafe_allow_html=True,
        )
        fig_dre = go.Figure(
            go.Waterfall(
                name="DRE",
                orientation="v",
                measure=["relative", "relative", "relative", "total"],
                x=[
                    "Faturamento Bruto",
                    "(-) Custos Variáveis",
                    "(-) Custos Fixos",
                    "(=) Lucro Líquido",
                ],
                textposition="outside",
                text=[
                    formatar_moeda(faturamento_bruto),
                    f"-{formatar_moeda(custos_variaveis)}",
                    f"-{formatar_moeda(custos_fixos)}",
                    formatar_moeda(lucro_liquido),
                ],
                y=[faturamento_bruto, -custos_variaveis, -custos_fixos, lucro_liquido],
                connector={"line": {"color": "#687787"}},
                increasing={"marker": {"color": "#1F9D55"}},
                decreasing={"marker": {"color": "#D9383A"}},
                totals={"marker": {"color": "#0B1B2B"}},
            )
        )
        fig_dre.update_layout(
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            font=dict(color="#0B1B2B", family="Poppins"),
            yaxis=dict(title="R$", gridcolor="#EBE4D8"),
            margin=dict(l=15, r=15, t=25, b=15),
            height=340,
        )
        st.plotly_chart(fig_dre, use_container_width=True)

    with g_col2:
        st.markdown(
            '<div class="mazon-card" style="margin-bottom:8px; padding:1.1rem 1.4rem 0.4rem 1.4rem;">'
            '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">📈 Fluxo de Caixa &amp; Faturamento por Período</h4>'
            "</div>",
            unsafe_allow_html=True,
        )
        if not df_filtrado.empty:
            df_fluxo = (
                df_filtrado.groupby(["Data", "Tipo"], as_index=False)["Valor"]
                .sum()
                .sort_values("Data")
            )
            fig_fluxo = px.bar(
                df_fluxo,
                x="Data",
                y="Valor",
                color="Tipo",
                barmode="group",
                color_discrete_map={"Receita": "#0B1B2B", "Despesa": "#D4A24C"},
            )
            fig_fluxo.update_layout(
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
                font=dict(color="#0B1B2B", family="Poppins"),
                xaxis=dict(title="Data", gridcolor="#EBE4D8"),
                yaxis=dict(title="Valor (R$)", gridcolor="#EBE4D8"),
                margin=dict(l=15, r=15, t=25, b=15),
                height=340,
            )
            st.plotly_chart(fig_fluxo, use_container_width=True)
        else:
            st.info("Sem lançamentos para o período selecionado.")

    g_col3, g_col4 = st.columns(2)
    with g_col3:
        st.markdown(
            '<div class="mazon-card" style="margin-bottom:8px; padding:1.1rem 1.4rem 0.4rem 1.4rem;">'
            '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">🍩 Distribuição de Despesas por Categoria</h4>'
            "</div>",
            unsafe_allow_html=True,
        )
        if not df_despesas.empty:
            df_desp_cat = df_despesas.groupby("Categoria", as_index=False)["Valor"].sum()
            fig_desp = px.pie(
                df_desp_cat,
                names="Categoria",
                values="Valor",
                hole=0.58,
                color_discrete_sequence=["#0B1B2B", "#D4A24C", "#1C354D", "#C59237", "#526375", "#D9383A"],
            )
            fig_desp.update_traces(textinfo="percent+label", textfont_color="#0B1B2B")
            fig_desp.update_layout(
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
                font=dict(color="#0B1B2B", family="Poppins"),
                showlegend=False,
                margin=dict(l=10, r=10, t=15, b=15),
                height=320,
            )
            st.plotly_chart(fig_desp, use_container_width=True)
        else:
            st.info("Nenhuma despesa nos filtros selecionados.")

    with g_col4:
        st.markdown(
            '<div class="mazon-card" style="margin-bottom:8px; padding:1.1rem 1.4rem 0.4rem 1.4rem;">'
            '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">💳 Faturamento por Forma de Pagamento</h4>'
            "</div>",
            unsafe_allow_html=True,
        )
        df_rec = df_filtrado[df_filtrado["Tipo"] == "Receita"]
        if not df_rec.empty:
            df_rec_fp = df_rec.groupby("Forma_Pagamento", as_index=False)["Valor"].sum()
            fig_fp = px.bar(
                df_rec_fp,
                x="Valor",
                y="Forma_Pagamento",
                orientation="h",
                text=df_rec_fp["Valor"].apply(formatar_moeda),
                color_discrete_sequence=["#0B1B2B"],
            )
            fig_fp.update_traces(textposition="outside")
            fig_fp.update_layout(
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
                font=dict(color="#0B1B2B", family="Poppins"),
                xaxis=dict(title="Valor (R$)", gridcolor="#EBE4D8"),
                yaxis=dict(title=""),
                margin=dict(l=10, r=90, t=15, b=15),
                height=320,
            )
            st.plotly_chart(fig_fp, use_container_width=True)
        else:
            st.info("Nenhuma receita nos filtros selecionados.")

    # 4. FORMULÁRIO DE NOVO LANÇAMENTO FINANCEIRO E UPLOAD DE NOTA FISCAL / COMPROVANTE
    with st.expander("➕ Novo Lançamento Financeiro / Upload de Nota Fiscal ou Comprovante (Google Drive)", expanded=False):
        with st.form("form_lancamento_completo"):
            c_l1, c_l2, c_l3, c_l4 = st.columns(4)
            with c_l1:
                l_data = st.date_input("Data", value=date.today())
                l_tipo = st.selectbox("Tipo *", options=["Receita", "Despesa"])
            with c_l2:
                l_natureza = st.selectbox(
                    "Natureza",
                    options=["Faturamento Bruto", "Custo Variável", "Custo Fixo"],
                )
                l_cat = st.selectbox("Categoria *", options=CATEGORIAS_FINANCEIRAS)
            with c_l3:
                l_forn = st.text_input("Fornecedor / Cliente", placeholder="Ex: Mar & Cia Imports")
                l_forma = st.selectbox("Forma de Pagamento", options=FORMAS_PAGAMENTO)
            with c_l4:
                l_valor = st.number_input("Valor (R$) *", min_value=0.01, value=100.0, step=10.0)
                l_subpasta = st.selectbox("Subpasta no Drive", options=["Notas_Fiscais", "Comprovantes"])

            l_desc = st.text_input("Descrição Detalhada *", placeholder="Ex: NF 882 Compra de Salmão ou Fechamento Jantar")
            l_arq = st.file_uploader(
                "Anexar Nota Fiscal ou Comprovante (.pdf, .png, .jpg) — Upload automático para a pasta principal no Google Drive",
                type=["pdf", "png", "jpg", "jpeg"],
            )

            if st.form_submit_button("Salvar Lançamento & Arquivar Documento"):
                link_doc = "-"
                if l_arq is not None:
                    link_doc = db.upload_arquivo_drive(
                        arquivo_bytes=l_arq.read(),
                        nome_arquivo=l_arq.name,
                        mime_type=l_arq.type or "application/octet-stream",
                        subpasta=l_subpasta,
                    )
                    registrar_log(
                        usuario=USUARIO_ATUAL,
                        perfil=PERFIL_ATUAL,
                        acao=f"Upload de arquivo '{l_arq.name}' para /{l_subpasta} no Google Drive",
                        modulo="Financeiro",
                    )

                ok_l, id_gerado = db.registrar_lancamento_financeiro(
                    data_lancamento=l_data.strftime("%Y-%m-%d"),
                    tipo=l_tipo,
                    natureza_custo=l_natureza,
                    descricao=l_desc,
                    categoria=l_cat,
                    fornecedor=l_forn,
                    forma_pagamento=l_forma,
                    valor=l_valor,
                    link_comprovante=link_doc,
                    responsavel=USUARIO_ATUAL,
                )
                if ok_l:
                    registrar_log(
                        usuario=USUARIO_ATUAL,
                        perfil=PERFIL_ATUAL,
                        acao=f"Novo lançamento financeiro ({id_gerado}) - {l_tipo}: {formatar_moeda(l_valor)} ({l_cat})",
                        modulo="Financeiro",
                    )
                    st.success(f"Lançamento {id_gerado} registrado com sucesso!")
                    st.rerun()

    # 5. TABELA DETALHADA DE LANÇAMENTOS E NOTAS FISCAIS CADASTRADAS
    st.markdown("#### 📋 Lançamentos e Notas Fiscais Cadastradas")
    linhas_fin_html = []
    for _, row in df_filtrado.sort_values("Data", ascending=False).iterrows():
        cor_tipo = "#1F9D55" if row["Tipo"] == "Receita" else "#D9383A"
        link_str = str(row.get("Link_Comprovante", "-"))
        link_html = (
            f'<a href="{link_str}" target="_blank" style="color:#9A6B1F; font-weight:600; text-decoration:underline;">Abrir no Drive</a>'
            if link_str.startswith("http")
            else '<span style="color:#687787;">—</span>'
        )
        linhas_fin_html.append(
            "<tr>"
            f'<td style="font-weight:600; color:#9A6B1F;">{row["ID_Lancamento"]}</td>'
            f'<td>{row["Data"]}</td>'
            f'<td style="font-weight:600; color:{cor_tipo};">{row["Tipo"]}</td>'
            f'<td>{row["Natureza_Custo"]}</td>'
            f'<td style="font-weight:500;">{row["Descricao"]}</td>'
            f'<td>{row["Categoria"]}</td>'
            f'<td>{row["Fornecedor"]}</td>'
            f'<td>{row["Forma_Pagamento"]}</td>'
            f'<td style="font-weight:600; color:{cor_tipo};">{formatar_moeda(float(row["Valor"]))}</td>'
            f"<td>{link_html}</td>"
            f'<td style="color:#687787;">{row["Responsavel"]}</td>'
            "</tr>"
        )

    corpo_fin = "".join(linhas_fin_html) if linhas_fin_html else (
        '<tr><td colspan="11" style="text-align:center; color:#687787; padding:18px;">Nenhum lançamento encontrado para os filtros aplicados.</td></tr>'
    )

    st.markdown(
        '<div class="mazon-table-wrapper">'
        '<table class="mazon-table">'
        "<thead>"
        "<tr>"
        "<th>ID</th>"
        "<th>Data</th>"
        "<th>Tipo</th>"
        "<th>Natureza</th>"
        "<th>Descrição</th>"
        "<th>Categoria</th>"
        "<th>Fornecedor</th>"
        "<th>Pagamento</th>"
        "<th>Valor</th>"
        "<th>Comprovante / NF</th>"
        "<th>Responsável</th>"
        "</tr>"
        "</thead>"
        f"<tbody>{corpo_fin}</tbody>"
        "</table>"
        "</div>",
        unsafe_allow_html=True,
    )

    # 6. TRAVA RIGOROSA DE EXCLUSÃO: EXCLUSIVAMENTE QUANDO st.session_state['usuario'] == 'Arthur'
    if st.session_state.get("usuario") == "Arthur":
        with st.expander("🗑️ Exclusão de Lançamentos Financeiros / Notas Fiscais (Exclusivo Arthur - Dev)", expanded=False):
            st.warning(
                "Atenção: Esta ação remove definitivamente o lançamento ou nota fiscal da planilha unificada e registra o evento no Log de Auditoria."
            )
            ids_disponiveis = df_fin["ID_Lancamento"].astype(str).tolist() if not df_fin.empty else []
            if ids_disponiveis:
                col_del1, col_del2 = st.columns([3, 1.2])
                with col_del1:
                    id_para_excluir = st.selectbox(
                        "Selecione o ID do Lançamento / Nota Fiscal para excluir:",
                        options=ids_disponiveis,
                        key="sel_delete_arthur",
                    )
                with col_del2:
                    st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                    if st.button("🗑️ Excluir Registro", key="btn_delete_arthur", use_container_width=True):
                        if st.session_state.get("usuario") == "Arthur":
                            ok_del, msg_del = db.excluir_lancamento_financeiro(
                                id_lancamento=id_para_excluir,
                                usuario_logado=USUARIO_ATUAL,
                                perfil_logado=PERFIL_ATUAL,
                            )
                            if ok_del:
                                st.success(msg_del)
                                st.rerun()
                            else:
                                st.error(msg_del)


def renderizar_tab_auditoria() -> None:
    """
    🛡️ Tab Auditoria:
    - Dev (Arthur): Acesso completo ao Log Geral (100% dos eventos).
    - Gerente (Claudio): Filtro automático exibindo apenas logs referentes aos módulos de Estoque e Caixa.
    """
    st.markdown(
        '<div style="margin-bottom: 16px;">'
        '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">🛡️ Logs de Auditoria &amp; Rastreabilidade</h1>'
        '<p style="color:#687787; font-size:0.9rem; margin:0;">Monitoramento contínuo de operações gravado automaticamente na aba <code>Logs_Auditoria</code> da planilha unificada.</p>'
        "</div>",
        unsafe_allow_html=True,
    )

    df_logs = db.ler_aba("Logs_Auditoria").copy()

    if PERFIL_ATUAL == "Gerente":
        st.info(
            "🔎 Visão Gerencial (Claudio): Exibindo automaticamente apenas os eventos referentes aos módulos de **Estoque** e **Caixa**."
        )
        df_logs = df_logs[df_logs["Modulo"].astype(str).isin(["Estoque", "Caixa"])].copy()
    else:
        st.markdown(
            '<div class="mazon-card" style="padding:0.9rem 1.3rem;">'
            '<span class="badge-best-price">ACESSO TOTAL DEV (ARTHUR)</span>'
            '<span style="margin-left:10px; font-size:0.86rem; color:#0B1B2B;">Visualizando 100% do Log Geral de Auditoria (Autenticação, Usuários, Estoque, Caixa e Financeiro).</span>'
            "</div>",
            unsafe_allow_html=True,
        )

    col_a1, col_a2 = st.columns(2)
    with col_a1:
        busca_log = st.text_input("Filtrar por Usuário ou Ação", placeholder="Ex: Arthur, Salmão, Login...")
    with col_a2:
        modulos_disp = sorted(df_logs["Modulo"].astype(str).unique().tolist()) if not df_logs.empty else []
        filtro_mod = st.multiselect("Filtrar por Módulo", options=modulos_disp, default=[])

    if busca_log.strip():
        termo_l = busca_log.strip().lower()
        df_logs = df_logs[
            df_logs["Usuario"].astype(str).str.lower().str.contains(termo_l)
            | df_logs["Acao"].astype(str).str.lower().str.contains(termo_l)
        ]
    if filtro_mod:
        df_logs = df_logs[df_logs["Modulo"].isin(filtro_mod)]

    df_logs = df_logs.iloc[::-1]  # Mais recentes no topo

    linhas_log_html = []
    for _, row in df_logs.iterrows():
        linhas_log_html.append(
            "<tr>"
            f'<td style="color:#687787; font-size:0.82rem;">{row["Data_Hora"]}</td>'
            f'<td style="font-weight:600; color:#0B1B2B;">{row["Usuario"]}</td>'
            f'<td><span class="badge-alerta">{row["Perfil"]}</span></td>'
            f'<td>{row["Acao"]}</td>'
            f'<td style="font-weight:600; color:#9A6B1F;">{row["Modulo"]}</td>'
            "</tr>"
        )

    corpo_logs = "".join(linhas_log_html) if linhas_log_html else (
        '<tr><td colspan="5" style="text-align:center; color:#687787; padding:18px;">Nenhum registro de auditoria encontrado.</td></tr>'
    )

    st.markdown(
        '<div class="mazon-table-wrapper">'
        '<table class="mazon-table">'
        "<thead>"
        "<tr>"
        "<th>Data / Hora</th>"
        "<th>Usuário</th>"
        "<th>Perfil</th>"
        "<th>Ação Executada</th>"
        "<th>Módulo</th>"
        "</tr>"
        "</thead>"
        f"<tbody>{corpo_logs}</tbody>"
        "</table>"
        "</div>",
        unsafe_allow_html=True,
    )


# ==============================================================================
# NAVEGAÇÃO POR TABS SUPERIORES DINÂMICAS (st.tabs) CONFORME O PERFIL LOGADO
# ==============================================================================
if PERFIL_ATUAL in ["Dev", "Gerente"]:
    tab_home, tab_estoque, tab_fin, tab_auditoria = st.tabs(
        ["🏠 Home", "📦 Estoque", "💰 Gestão Financeira", "🛡️ Auditoria"]
    )
    with tab_home:
        renderizar_tab_home()
    with tab_estoque:
        renderizar_tab_estoque()
    with tab_fin:
        renderizar_tab_financeiro()
    with tab_auditoria:
        renderizar_tab_auditoria()

elif PERFIL_ATUAL == "Caixa":
    tab_home, tab_estoque, tab_caixa = st.tabs(
        ["🏠 Home", "📦 Estoque", "💰 Lançamentos / Caixa"]
    )
    with tab_home:
        renderizar_tab_home()
    with tab_estoque:
        renderizar_tab_estoque()
    with tab_caixa:
        renderizar_tab_financeiro()

else:  # Perfil Equipe
    tab_home, tab_estoque = st.tabs(["🏠 Home", "📦 Estoque"])
    with tab_home:
        renderizar_tab_home()
    with tab_estoque:
        renderizar_tab_estoque()
