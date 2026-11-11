"""
Restaurante Mazon — Sistema Unificado de Gestão de Estoque, Inteligência Financeira & RBAC
Arquitetura:
- Navegação em Cascata: Tabs Superiores + Sidebar Contextual Dinâmica
- Banco Unificado em 1 Planilha no Google Sheets (Abas: Usuarios, Logs_Auditoria, Estoque, Financeiro)
- Repositório Único no Google Drive com subpastas /Notas_Fiscais e /Comprovantes
- Controle de Acesso por Hierarquia (Dev > Gerente > Caixa > Equipe) + Matriz de Permissões Granulares
- Trava Anti-Fraude de Exclusão (Exclusiva para Arthur - Dev)
- Sistema Multiformato de Exportação Nativa (CSV, XLSX e PDF via ReportLab)
"""

from __future__ import annotations

import io
import os
import re
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
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
    PERMISSOES_DISPONIVEIS,
    PERMISSOES_POR_PERFIL,
    MazonSheetsManager,
    registrar_log,
)

# ==============================================================================
# CONFIGURAÇÃO INICIAL DA PÁGINA (FAVICON ESTRITO logo2.png)
# ==============================================================================
icone_aba = Image.open("logo2.png")

st.set_page_config(
    page_title="Restaurante Mazon | Gestão de Estoque & Inteligência Financeira",
    page_icon=icone_aba,
    layout="wide",
    initial_sidebar_state="expanded",
)

# CSS baseweb para ocultar círculos de radio na barra lateral e no menu superior
st.markdown(
    """
<style>
/* Esconde totalmente a bolinha/círculo do st.radio (todas as versões do Streamlit) */
div[data-testid="stRadioButton"] [role="radio"],
div[data-testid="stRadioButton"] label > div:first-child,
div[data-baseweb="radio"] > div:first-child {
    display: none !important;
    visibility: hidden !important;
    width: 0px !important;
    height: 0px !important;
}

/* Remaneja o alinhamento do texto */
div[data-testid="stRadioButton"] label {
    padding-left: 5px !important;
    margin-left: 0px !important;
    cursor: pointer !important;
}

/* Destaca o item selecionado */
div[data-testid="stRadioButton"] label:has(input:checked) {
    font-weight: bold !important;
}
</style>
""",
    unsafe_allow_html=True,
)


def obter_caminho_logo() -> Optional[str]:
    """Busca dinamicamente o arquivo da logo entre os formatos suportados."""
    formatos = [
        "logo.jpeg",
        "logo.jpg",
        "logo.png",
        "WhatsApp Image 2026-11-09 at 17.04.39.jpeg",
    ]
    for fmt in formatos:
        if os.path.exists(fmt):
            return fmt
    return None


def carregar_estilos_css(caminho_css: str = "style.css") -> None:
    """Injeta a folha de estilos Luxury Minimalist UI (Creme #F4EDE3 & Azul Marinho #0B1B2B)."""
    if os.path.exists(caminho_css):
        with open(caminho_css, "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)


carregar_estilos_css()

# Inicializa o Gerenciador Unificado (Sheets + Drive + RBAC + Auditoria)
db = MazonSheetsManager()


# ==============================================================================
# UTILITÁRIOS VISUAIS E MOTOR INTEGRADO DE EXPORTAÇÃO (CSV, XLSX E PDF REPORTLAB)
# ==============================================================================
def formatar_moeda(valor: float) -> str:
    """Formata um float no padrão monetário brasileiro (R$ 1.234,56)."""
    try:
        v = float(valor)
    except (ValueError, TypeError):
        v = 0.0
    sinal = "-" if v < 0 else ""
    v_abs = abs(v)
    texto = f"{v_abs:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sinal}R$ {texto}"


def renderizar_badge_status(status: str) -> str:
    """Retorna a pill HTML estilizada para o status do estoque."""
    st_upper = str(status).upper()
    if "CRÍTICO" in st_upper or "FALTA" in st_upper:
        return '<span class="badge-critico">Crítico / Falta</span>'
    elif "ALERTA" in st_upper:
        return '<span class="badge-alerta">Alerta Reposição</span>'
    elif "PENDENTE" in st_upper:
        return '<span class="badge-pendente">Pendente Contagem</span>'
    return '<span class="badge-ok">Estoque OK</span>'


def gerar_csv_bytes(df: pd.DataFrame) -> bytes:
    """Gera bytes CSV compatíveis com Excel PT-BR (UTF-8 com BOM e separador ';')."""
    return df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")


def gerar_xlsx_bytes(df: pd.DataFrame, nome_aba: str = "Relatorio_Mazon") -> bytes:
    """Gera arquivo Excel nativo (.xlsx) em memória utilizando openpyxl."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        aba_limpa = re.sub(r"[\\/*?:\[\]]", "_", nome_aba)[:31] or "Dados"
        df.to_excel(writer, index=False, sheet_name=aba_limpa)
    buffer.seek(0)
    return buffer.getvalue()


def _decorar_pagina_pdf(canvas: Any, doc: Any) -> None:
    """Desenha cabeçalho/rodapé institucional Luxury Navy & Gold nas páginas do PDF."""
    canvas.saveState()
    largura, altura = doc.pagesize

    # Faixa superior Azul Marinho e filete Dourado
    canvas.setFillColor(colors.HexColor("#0B1B2B"))
    canvas.rect(0, altura - 10 * mm, largura, 10 * mm, fill=1, stroke=0)
    canvas.setFillColor(colors.HexColor("#D4A24C"))
    canvas.rect(0, altura - 11.2 * mm, largura, 1.2 * mm, fill=1, stroke=0)

    # Rodapé com numeração e assinatura do sistema
    canvas.setStrokeColor(colors.HexColor("#E8DFD2"))
    canvas.setLineWidth(0.6)
    canvas.line(12 * mm, 12 * mm, largura - 12 * mm, 12 * mm)

    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#687787"))
    canvas.drawString(
        12 * mm,
        7 * mm,
        "Restaurante Mazon · Sistema Unificado de Estoque, Inteligência Financeira & Auditoria",
    )
    canvas.drawRightString(
        largura - 12 * mm,
        7 * mm,
        f"Página {doc.page} · Emitido em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
    )
    canvas.restoreState()


def gerar_pdf_tabela_analitica(
    df: pd.DataFrame,
    titulo: str,
    subtitulo: str = "",
) -> bytes:
    """
    Gera um relatório PDF com tabela formatada via ReportLab para qualquer tabela analítica
    (Estoque, Lançamentos Financeiros, Log de Auditoria, Comparativo ou Lista de Compras).
    """
    buffer = io.BytesIO()
    usar_paisagem = len(df.columns) > 6
    page_size = landscape(A4) if usar_paisagem else A4

    doc = SimpleDocTemplate(
        buffer,
        pagesize=page_size,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
    )

    estilos = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "MazonTituloTabela",
        parent=estilos["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=18,
        alignment=1,
        textColor=colors.HexColor("#0B1B2B"),
        spaceAfter=4,
    )
    estilo_sub = ParagraphStyle(
        "MazonSubTabela",
        parent=estilos["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        alignment=1,
        textColor=colors.HexColor("#687787"),
        spaceAfter=10,
    )
    estilo_celula = ParagraphStyle(
        "MazonCelula",
        parent=estilos["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor("#0B1B2B"),
    )
    estilo_cabecalho = ParagraphStyle(
        "MazonCabecalho",
        parent=estilos["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.8,
        leading=10,
        textColor=colors.HexColor("#F4EDE3"),
    )

    elementos: List[Any] = []
    caminho_logo = obter_caminho_logo()
    if caminho_logo and os.path.exists(caminho_logo):
        try:
            img_logo = RLImage(caminho_logo, width=110, height=70)
            img_logo.hAlign = "CENTER"
            elementos.append(img_logo)
            elementos.append(Spacer(1, 3 * mm))
        except Exception:
            pass

    elementos.append(Paragraph(f"RESTAURANTE MAZON — {titulo.upper()}", estilo_titulo))
    info_sub = subtitulo or f"Total de registros exportados: {len(df)} | Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    elementos.append(Paragraph(info_sub, estilo_sub))

    if df.empty:
        elementos.append(
            Paragraph(
                "Nenhum registro encontrado para os filtros selecionados (Tabela vazia / R$ 0,00).",
                estilo_sub,
            )
        )
    else:
        colunas = [str(c) for c in df.columns]
        linhas_dados: List[List[Any]] = [[Paragraph(c, estilo_cabecalho) for c in colunas]]

        for _, row in df.iterrows():
            linha_formatada = []
            for col in df.columns:
                val = row[col]
                if isinstance(val, float):
                    if any(k in col.lower() for k in ["valor", "preco", "preço", "custo", "total"]):
                        txt = formatar_moeda(val)
                    else:
                        txt = f"{val:.2f}"
                else:
                    txt = str(val) if pd.notna(val) else "-"
                linha_formatada.append(Paragraph(txt[:90], estilo_celula))
            linhas_dados.append(linha_formatada)

        largura_util = page_size[0] - 24 * mm
        largura_col = largura_util / max(len(colunas), 1)
        tabela = Table(linhas_dados, colWidths=[largura_col] * len(colunas), repeatRows=1)
        tabela.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B1B2B")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#F4EDE3")),
                    ("LINEBELOW", (0, 0), (-1, 0), 1.5, colors.HexColor("#D4A24C")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F4EDE3")]),
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#E5DCD0")),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )
        elementos.append(tabela)

    doc.build(elementos, onFirstPage=_decorar_pagina_pdf, onLaterPages=_decorar_pagina_pdf)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


def gerar_pdf_relatorio_resumido(
    titulo: str,
    periodo_str: str,
    kpis: List[Tuple[str, str, str]],
    secoes_texto: Optional[List[Tuple[str, str]]] = None,
    df_resumo: Optional[pd.DataFrame] = None,
) -> bytes:
    """
    Gera um Relatório Executivo Resumido em PDF (para Home, Estoque Dashboard, DRE, Faturamento,
    CMV, Break-Even e Relatório Descritivo) contendo consolidação dos KPIs, período e tabela resumo.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
    )

    estilos = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "MazonExecTitulo",
        parent=estilos["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=19,
        alignment=1,
        textColor=colors.HexColor("#0B1B2B"),
        spaceAfter=4,
    )
    estilo_sub = ParagraphStyle(
        "MazonExecSub",
        parent=estilos["Normal"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=13,
        alignment=1,
        textColor=colors.HexColor("#687787"),
        spaceAfter=14,
    )
    estilo_secao = ParagraphStyle(
        "MazonExecSecao",
        parent=estilos["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=15,
        textColor=colors.HexColor("#0B1B2B"),
        spaceBefore=10,
        spaceAfter=6,
    )
    estilo_corpo = ParagraphStyle(
        "MazonExecCorpo",
        parent=estilos["Normal"],
        fontName="Helvetica",
        fontSize=9.2,
        leading=13.5,
        textColor=colors.HexColor("#0B1B2B"),
        spaceAfter=6,
    )

    elementos: List[Any] = []
    caminho_logo = obter_caminho_logo()
    if caminho_logo and os.path.exists(caminho_logo):
        try:
            img_logo = RLImage(caminho_logo, width=120, height=78)
            img_logo.hAlign = "CENTER"
            elementos.append(img_logo)
            elementos.append(Spacer(1, 3 * mm))
        except Exception:
            pass

    elementos.append(Paragraph(f"RESTAURANTE MAZON — {titulo.upper()}", estilo_titulo))
    elementos.append(
        Paragraph(
            f"<b>Período / Referência:</b> {periodo_str} &nbsp;|&nbsp; "
            f"<b>Emissão:</b> {datetime.now().strftime('%d/%m/%Y às %H:%M')}",
            estilo_sub,
        )
    )

    # Tabela Consolidada de KPIs
    if kpis:
        elementos.append(Paragraph("1. Consolidação dos Indicadores-Chave (KPIs)", estilo_secao))
        dados_kpi = [["Indicador / Métrica", "Valor Consolidado", "Detalhamento / Observação"]]
        for nome_kpi, valor_kpi, obs_kpi in kpis:
            dados_kpi.append([str(nome_kpi), str(valor_kpi), str(obs_kpi)])

        tab_kpi = Table(dados_kpi, colWidths=[60 * mm, 45 * mm, 77 * mm], repeatRows=1)
        tab_kpi.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B1B2B")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#F4EDE3")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 9),
                    ("LINEBELOW", (0, 0), (-1, 0), 1.5, colors.HexColor("#D4A24C")),
                    ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
                    ("FONTNAME", (1, 1), (1, -1), "Helvetica-Bold"),
                    ("TEXTCOLOR", (1, 1), (1, -1), colors.HexColor("#9A6B1F")),
                    ("FONTNAME", (2, 1), (2, -1), "Helvetica"),
                    ("FONTSIZE", (0, 1), (-1, -1), 8.8),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F4EDE3")]),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E5DCD0")),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        elementos.append(tab_kpi)
        elementos.append(Spacer(1, 4 * mm))

    # Seções descritivas
    if secoes_texto:
        for tit_sec, corpo_sec in secoes_texto:
            elementos.append(Paragraph(tit_sec, estilo_secao))
            elementos.append(Paragraph(corpo_sec, estilo_corpo))

    # Tabela Resumo opcional
    if df_resumo is not None and not df_resumo.empty:
        elementos.append(Paragraph("Detalhamento Analítico do Período", estilo_secao))
        cols = [str(c) for c in df_resumo.columns]
        dados_tab = [cols]
        for _, r in df_resumo.head(60).iterrows():
            linha = []
            for c in df_resumo.columns:
                v = r[c]
                if isinstance(v, float):
                    linha.append(formatar_moeda(v) if any(k in c.lower() for k in ["valor", "preco", "custo", "total", "saldo", "resultado"]) else f"{v:.2f}")
                else:
                    linha.append(str(v)[:45])
            dados_tab.append(linha)

        larg_col = (182 * mm) / max(len(cols), 1)
        t_res = Table(dados_tab, colWidths=[larg_col] * len(cols), repeatRows=1)
        t_res.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B1B2B")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#F4EDE3")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F4EDE3")]),
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#E5DCD0")),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )
        elementos.append(t_res)

    doc.build(elementos, onFirstPage=_decorar_pagina_pdf, onLaterPages=_decorar_pagina_pdf)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


def renderizar_barra_exportacao_tabela(
    df: pd.DataFrame,
    prefixo_arquivo: str,
    titulo_pdf: str,
    subtitulo_pdf: str,
    key_suffix: str,
) -> None:
    """Renderiza os 3 botões padronizados de exportação (CSV, XLSX e PDF) para tabelas analíticas."""
    data_tag = datetime.now().strftime("%Y%m%d_%H%M")
    csv_data = gerar_csv_bytes(df)
    xlsx_data = gerar_xlsx_bytes(df, nome_aba=prefixo_arquivo)
    pdf_data = gerar_pdf_tabela_analitica(df, titulo=titulo_pdf, subtitulo=subtitulo_pdf)

    c_exp1, c_exp2, c_exp3 = st.columns(3)
    with c_exp1:
        st.download_button(
            label="📄 Exportar CSV",
            data=csv_data,
            file_name=f"{prefixo_arquivo}_{data_tag}.csv",
            mime="text/csv",
            use_container_width=True,
            key=f"btn_csv_{key_suffix}",
        )
    with c_exp2:
        st.download_button(
            label="📊 Exportar XLSX (Excel)",
            data=xlsx_data,
            file_name=f"{prefixo_arquivo}_{data_tag}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"btn_xlsx_{key_suffix}",
        )
    with c_exp3:
        st.download_button(
            label="📑 Exportar Tabela em PDF",
            data=pdf_data,
            file_name=f"{prefixo_arquivo}_{data_tag}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key=f"btn_pdf_{key_suffix}",
        )


# ==============================================================================
# ESTADO DE AUTENTICAÇÃO E TELA DE LOGIN INICIAL (PLACEHOLDER NEUTRO)
# ==============================================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "usuario" not in st.session_state:
    st.session_state["usuario"] = None
if "perfil" not in st.session_state:
    st.session_state["perfil"] = None
if "permissoes" not in st.session_state:
    st.session_state["permissoes"] = []


def renderizar_tela_login() -> None:
    """Renderiza a tela central de login com campo neutro (sem dicas de nomes de usuários)."""
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
            in_usuario = st.text_input("Usuário", placeholder="Digite seu usuário")
            in_senha = st.text_input("Senha", type="password", placeholder="Digite sua senha")
            btn_entrar = st.form_submit_button("Entrar no Sistema", use_container_width=True)

            if btn_entrar:
                dados_auth = db.autenticar_usuario(in_usuario, in_senha)
                if dados_auth:
                    st.session_state["autenticado"] = True
                    st.session_state["usuario"] = dados_auth["usuario"]
                    st.session_state["perfil"] = dados_auth["perfil"]
                    st.session_state["permissoes"] = dados_auth.get(
                        "permissoes",
                        list(PERMISSOES_POR_PERFIL.get(dados_auth["perfil"], ["home", "estoque_leitura"])),
                    )
                    registrar_log(
                        usuario=dados_auth["usuario"],
                        perfil=dados_auth["perfil"],
                        acao=f"Login realizado com sucesso ({dados_auth['perfil']})",
                        modulo="Autenticacao",
                    )
                    st.rerun()
                else:
                    registrar_log(
                        usuario=in_usuario.strip() or "Desconhecido",
                        perfil="Nao_Autenticado",
                        acao="Tentativa falha de autenticação",
                        modulo="Autenticacao",
                    )
                    st.error("Usuário ou senha inválidos. Verifique suas credenciais.")


if not st.session_state["autenticado"]:
    renderizar_tela_login()
    st.stop()


# ==============================================================================
# SESSÃO AUTENTICADA: VARIÁVEIS DE PERFIL, PERMISSÕES E CABEÇALHO DA SIDEBAR
# ==============================================================================
USUARIO_ATUAL: str = str(st.session_state["usuario"])
PERFIL_ATUAL: str = str(st.session_state["perfil"])
PERMISSOES_ATUAIS: List[str] = list(
    st.session_state.get(
        "permissoes",
        PERMISSOES_POR_PERFIL.get(PERFIL_ATUAL, ["home", "estoque_leitura"]),
    )
)


def tem_permissao(chave: str) -> bool:
    """Verifica se o usuário logado possui determinada permissão granular."""
    if PERFIL_ATUAL == "Dev" and USUARIO_ATUAL == "Arthur":
        return True
    return chave in PERMISSOES_ATUAIS


# Renderiza a Logo no topo da Sidebar
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

# Construção dinâmica das Tabs Superiores disponíveis conforme perfil e permissões
TABS_DISPONIVEIS: List[str] = ["🏠 Home"]
if tem_permissao("estoque_leitura") or tem_permissao("estoque_edicao"):
    TABS_DISPONIVEIS.append("📦 Estoque")
if tem_permissao("fin_lancamentos") or tem_permissao("fin_relatorios") or PERFIL_ATUAL in ["Dev", "Gerente", "Caixa"]:
    TABS_DISPONIVEIS.append("💰 Gestão Financeira")
if (
    PERFIL_ATUAL in ["Dev", "Gerente"]
    or tem_permissao("auditoria_logs")
    or tem_permissao("usuarios_gestao")
):
    TABS_DISPONIVEIS.append("🛡️ Gestão de Usuários & Auditoria")

# Barra de Navegação Superior (Tabs Superiores em Cascata)
ABA_SUPERIOR_ATIVA = st.radio(
    "Módulo Principal",
    options=TABS_DISPONIVEIS,
    horizontal=True,
    label_visibility="collapsed",
    key="nav_top_tab_ativa",
)

# ==============================================================================
# SIDEBAR CONTEXTUAL DINÂMICA (ATUALIZA CONFORME A TAB SUPERIOR SELECIONADA)
# ==============================================================================
SUBPAGINA_ESTOQUE = "› Dashboard Principal"
SUBPAGINA_FINANCEIRO = "› Lançamentos e NFs"
SUBPAGINA_AUDITORIA = "› Log de Auditoria"

with st.sidebar:
    # Cartão de Identificação do Perfil Logado (Presente em todas as abas)
    st.markdown(
        f'<div style="background: rgba(255,255,255,0.05); border: 1px solid rgba(212,162,76,0.35); border-radius: 10px; padding: 12px 14px; margin-bottom: 10px;">'
        f'<div style="font-size: 0.72rem; text-transform: uppercase; letter-spacing: 1px; color: #D4A24C;">Sessão Ativa · Perfil {PERFIL_ATUAL}</div>'
        f'<div style="font-size: 1.02rem; font-weight: 600; color: #F4EDE3; margin-top: 2px;">👤 {USUARIO_ATUAL}</div>'
        "</div>",
        unsafe_allow_html=True,
    )

    # Contexto 1: Tab Home -> Exibe apenas dados do perfil logado e botão Sair
    if ABA_SUPERIOR_ATIVA == "🏠 Home":
        cor_status_nuvem = "#1F9D55" if db.connected else "#D4A24C"
        titulo_nuvem = "Mazon - Banco de Dados Online" if db.connected else "Banco Unificado Ativo"
        st.markdown(
            f'<div style="padding: 8px 12px; margin-bottom: 12px; border-left: 2px solid {cor_status_nuvem};">'
            f'<div style="font-size: 0.74rem; font-weight: 500; color: {cor_status_nuvem};">● {titulo_nuvem}</div>'
            f'<div style="font-size: 0.71rem; color: #C5D1DE; opacity: 0.78; margin-top: 2px;">{db.connection_message}</div>'
            "</div>",
            unsafe_allow_html=True,
        )

    # Contexto 2: Tab Estoque -> Exibe as 5 sub-páginas exclusivas de Estoque
    elif ABA_SUPERIOR_ATIVA == "📦 Estoque":
        st.markdown(
            '<p style="font-size:0.7rem; text-transform:uppercase; letter-spacing:1.4px; color:#D4A24C; font-weight:600; margin: 6px 0 6px 4px;">📦 Sub-páginas de Estoque</p>',
            unsafe_allow_html=True,
        )
        SUBPAGINA_ESTOQUE = st.radio(
            "Navegação de Estoque",
            options=[
                "› Dashboard Principal",
                "› Gestão e Tabela de Estoque",
                "› Comparativo de Preços",
                "› Central OCR / IA",
                "› Gerador de Lista de Compras",
            ],
            label_visibility="collapsed",
            key="sb_radio_estoque",
        )

    # Contexto 3: Tab Gestão Financeira -> Exibe as 6 sub-páginas (ou apenas Lançamentos e NFs para Caixa)
    elif ABA_SUPERIOR_ATIVA == "💰 Gestão Financeira":
        st.markdown(
            '<p style="font-size:0.7rem; text-transform:uppercase; letter-spacing:1.4px; color:#D4A24C; font-weight:600; margin: 6px 0 6px 4px;">💰 Sub-páginas Financeiras</p>',
            unsafe_allow_html=True,
        )
        if PERFIL_ATUAL == "Caixa" and not tem_permissao("fin_relatorios"):
            opcoes_fin_sb = ["› Lançamentos e NFs"]
        else:
            opcoes_fin_sb = [
                "› Lançamentos e NFs",
                "› Faturamento",
                "› DRE (Demonstrativo do Resultado)",
                "› CMV (Custo de Mercadoria Vendida)",
                "› Break-Even (Ponto de Equilíbrio)",
                "› Relatório Descritivo",
            ]
        SUBPAGINA_FINANCEIRO = st.radio(
            "Navegação Financeira",
            options=opcoes_fin_sb,
            label_visibility="collapsed",
            key="sb_radio_financeiro",
        )

    # Contexto 4: Tab Gestão de Usuários & Auditoria -> Exibe as 2 sub-páginas
    elif ABA_SUPERIOR_ATIVA == "🛡️ Gestão de Usuários & Auditoria":
        st.markdown(
            '<p style="font-size:0.7rem; text-transform:uppercase; letter-spacing:1.4px; color:#D4A24C; font-weight:600; margin: 6px 0 6px 4px;">🛡️ Segurança &amp; Governança</p>',
            unsafe_allow_html=True,
        )
        opcoes_aud_sb = []
        if PERFIL_ATUAL in ["Dev", "Gerente"] or tem_permissao("auditoria_logs"):
            opcoes_aud_sb.append("› Log de Auditoria")
        if PERFIL_ATUAL in ["Dev", "Gerente"] or tem_permissao("usuarios_gestao"):
            opcoes_aud_sb.append("› Gestão de Usuários & Permissões")
        if not opcoes_aud_sb:
            opcoes_aud_sb = ["› Log de Auditoria"]

        SUBPAGINA_AUDITORIA = st.radio(
            "Navegação de Auditoria",
            options=opcoes_aud_sb,
            label_visibility="collapsed",
            key="sb_radio_auditoria",
        )

    st.markdown('<hr style="border-color: rgba(255,255,255,0.08); margin: 14px 0;">', unsafe_allow_html=True)

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
        st.session_state["permissoes"] = []
        st.rerun()


# ==============================================================================
# MÓDULO 1: TAB 🏠 HOME (INDICADORES E ALERTAS OPERACIONAIS/FINANCEIROS)
# ==============================================================================
def renderizar_modulo_home() -> None:
    """Renderiza a página central da Tab Home com exportação de Relatório Resumido em PDF."""
    hoje_str = datetime.now().strftime("%Y-%m-%d")
    df_estoque = db.obter_estoque_enriquecido()
    df_criticos = df_estoque[df_estoque["Status"].isin(["CRÍTICO/FALTA", "ALERTA"])].copy()
    itens_contados = int((df_estoque["Qtd_Atual"] > 0).sum()) if not df_estoque.empty else 0
    valor_est = float(df_estoque["Valor_Total_Item"].sum()) if not df_estoque.empty else 0.0

    df_fin = db.ler_aba("Financeiro").copy()
    df_fin["Valor"] = pd.to_numeric(df_fin["Valor"], errors="coerce").fillna(0.0)
    df_hoje = df_fin[df_fin["Data"].astype(str) == hoje_str]
    fat_hoje = float(df_hoje[df_hoje["Tipo"] == "Receita"]["Valor"].sum()) if not df_hoje.empty else 0.0
    desp_hoje = float(df_hoje[df_hoje["Tipo"] == "Despesa"]["Valor"].sum()) if not df_hoje.empty else 0.0
    fat_acumulado = float(df_fin[df_fin["Tipo"] == "Receita"]["Valor"].sum()) if not df_fin.empty else 0.0
    desp_acumulada = float(df_fin[df_fin["Tipo"] == "Despesa"]["Valor"].sum()) if not df_fin.empty else 0.0
    saldo_liquido = fat_acumulado - desp_acumulada

    col_h_tit, col_h_pdf = st.columns([3.2, 1.2])
    with col_h_tit:
        st.markdown(
            f'<div style="margin-bottom: 14px;">'
            f'<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">🏠 Home — Visão Executiva &amp; Alertas Críticos</h1>'
            f'<p style="color:#687787; font-size:0.9rem; margin:0;">Bem-vindo, <strong>{USUARIO_ATUAL}</strong> ({PERFIL_ATUAL}) · Síntese operacional de {datetime.now().strftime("%d/%m/%Y")}</p>'
            "</div>",
            unsafe_allow_html=True,
        )
    with col_h_pdf:
        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
        kpis_home_pdf = [
            ("Faturamento de Hoje", formatar_moeda(fat_hoje), f"Despesas de Hoje: {formatar_moeda(desp_hoje)}"),
            ("Resultado Líquido Acumulado", formatar_moeda(saldo_liquido), f"Receitas Totais: {formatar_moeda(fat_acumulado)}"),
            ("Insumos em Alerta Crítico", str(len(df_criticos)), f"Contagens ativas em estoque: {itens_contados}"),
            ("Capital Imobilizado em Estoque", formatar_moeda(valor_est), "Avaliado pelo custo unitário cadastrado"),
        ]
        pdf_home_bytes = gerar_pdf_relatorio_resumido(
            titulo="Relatório Executivo — Home & Alertas",
            periodo_str=datetime.now().strftime("%d/%m/%Y"),
            kpis=kpis_home_pdf,
            secoes_texto=[
                (
                    "2. Diagnóstico Operacional do Dia",
                    f"O sistema registra {itens_contados} insumos com contagem ativa, {len(df_criticos)} alertas críticos de reposição "
                    f"e saldo líquido acumulado de {formatar_moeda(saldo_liquido)}.",
                )
            ],
            df_resumo=df_criticos[["ID", "Item", "Categoria", "Qtd_Atual", "Est_Minimo", "Status"]] if not df_criticos.empty else None,
        )
        st.download_button(
            label="📑 Exportar Relatório Resumido (PDF)",
            data=pdf_home_bytes,
            file_name=f"relatorio_home_mazon_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="btn_pdf_home_resumo",
        )

    if PERFIL_ATUAL in ["Dev", "Gerente"] or tem_permissao("fin_relatorios"):
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
                '<div class="mazon-kpi-label">Insumos em Alerta Crítico</div>'
                f'<div class="mazon-kpi-value" style="color:#D9383A;">{len(df_criticos)}</div>'
                f'<div class="mazon-kpi-sub">Contagens ativas em estoque: {itens_contados}</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with c4:
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
            if not df_fin.empty and float(df_fin["Valor"].sum()) > 0.0:
                df_resumo_dia = df_fin.groupby(["Data", "Tipo"], as_index=False)["Valor"].sum()
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
                    height=320,
                )
                st.plotly_chart(fig_home, use_container_width=True)
            else:
                st.info("Sem lançamentos financeiros registrados na planilha (R$ 0,00).")

        with col_g2:
            st.markdown(
                '<div class="mazon-card" style="margin-bottom: 8px; padding: 1.15rem 1.4rem 0.4rem 1.4rem;">'
                '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">Saldo Físico Contado por Categoria</h4>'
                "</div>",
                unsafe_allow_html=True,
            )
            df_com_saldo = df_estoque[df_estoque["Qtd_Atual"] > 0].copy()
            if not df_com_saldo.empty:
                df_contagem_cat = (
                    df_com_saldo.groupby("Categoria", as_index=False)["Qtd_Atual"]
                    .sum()
                    .rename(columns={"Qtd_Atual": "Saldo_Total"})
                )
                fig_pie = px.pie(
                    df_contagem_cat,
                    names="Categoria",
                    values="Saldo_Total",
                    hole=0.6,
                    color_discrete_sequence=["#0B1B2B", "#D4A24C", "#1C354D", "#C59237", "#34506B", "#E0B86B"],
                )
                fig_pie.update_traces(textinfo="percent+label", textfont_color="#0B1B2B")
                fig_pie.update_layout(
                    paper_bgcolor="#FFFFFF",
                    plot_bgcolor="#FFFFFF",
                    font=dict(color="#0B1B2B", family="Poppins"),
                    showlegend=False,
                    margin=dict(l=10, r=10, t=10, b=10),
                    height=320,
                )
                st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.info("Sem contagens de estoque registradas na planilha (0 itens contados).")

    elif PERFIL_ATUAL == "Caixa":
        lancamentos_hoje_caixa = df_fin[
            (df_fin["Data"].astype(str) == hoje_str)
            & (df_fin["Responsavel"].astype(str) == USUARIO_ATUAL)
        ]
        total_hoje_caixa = float(lancamentos_hoje_caixa["Valor"].sum()) if not lancamentos_hoje_caixa.empty else 0.0

        c_cx1, c_cx2, c_cx3 = st.columns(3)
        with c_cx1:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Seus Lançamentos Hoje</div>'
                f'<div class="mazon-kpi-value">{len(lancamentos_hoje_caixa)}</div>'
                f'<div class="mazon-kpi-sub">Total registrado hoje: {formatar_moeda(total_hoje_caixa)}</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with c_cx2:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Rotina 1 · Fechamento Diário</div>'
                '<div class="mazon-kpi-value" style="font-size:1.2rem;">💰 Gestão Financeira</div>'
                '<div class="mazon-kpi-sub">Acesse Lançamentos e NFs para informar o Faturamento Bruto</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with c_cx3:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Rotina 2 · Notas &amp; Comprovantes</div>'
                '<div class="mazon-kpi-value" style="font-size:1.2rem;">📎 Upload Google Drive</div>'
                '<div class="mazon-kpi-sub">Envie PDFs e fotos de NFs para /Notas_Fiscais e /Comprovantes</div>'
                "</div>",
                unsafe_allow_html=True,
            )

    else:
        c_eq1, c_eq2, c_eq3 = st.columns(3)
        with c_eq1:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Insumos em Alerta Crítico</div>'
                f'<div class="mazon-kpi-value" style="color:#D9383A;">{len(df_criticos)}</div>'
                '<div class="mazon-kpi-sub">Itens abaixo do estoque mínimo definido</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with c_eq2:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Contagens Ativas Registradas</div>'
                f'<div class="mazon-kpi-value">{itens_contados}</div>'
                '<div class="mazon-kpi-sub">Itens com saldo físico maior que zero</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with c_eq3:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Valor Estimado em Estoque</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(valor_est)}</div>'
                '<div class="mazon-kpi-sub">Soma das quantidades × custo unitário</div>'
                "</div>",
                unsafe_allow_html=True,
            )

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
        st.markdown(
            '<div class="critical-hero-card" style="margin-top: 14px;">'
            '<div class="critical-hero-header">'
            f'<div class="critical-hero-title">⚠️ Alertas Operacionais de Estoque Crítico ({len(df_criticos)} itens)</div>'
            '<span class="badge-best-price">Reposição Prioritária</span>'
            "</div>"
            f'<div class="critical-items-grid">{"".join(pills_list)}</div>'
            "</div>",
            unsafe_allow_html=True,
        )


# ==============================================================================
# MÓDULO 2: TAB 📦 ESTOQUE (5 SUB-PÁGINAS NA SIDEBAR CONTEXTUAL)
# ==============================================================================
def renderizar_modulo_estoque(subpagina: str) -> None:
    """Renderiza as 5 sub-páginas exclusivas da Tab Estoque conforme seleção na Sidebar."""
    df_estoque = db.obter_estoque_enriquecido()
    pode_editar = PERFIL_ATUAL in ["Dev", "Gerente"] or tem_permissao("estoque_edicao")

    # --------------------------------------------------------------------------
    # SUB-PÁGINA 1: DASHBOARD PRINCIPAL DE ESTOQUE
    # --------------------------------------------------------------------------
    if subpagina == "› Dashboard Principal":
        df_crit = df_estoque[df_estoque["Status"].isin(["CRÍTICO/FALTA", "ALERTA"])].copy()
        itens_contados = int((df_estoque["Qtd_Atual"] > 0).sum()) if not df_estoque.empty else 0
        capital_estoque = float(df_estoque["Valor_Total_Item"].sum()) if not df_estoque.empty else 0.0
        itens_ok = int((df_estoque["Status"] == "OK/CHEIO").sum()) if not df_estoque.empty else 0

        c_t1, c_t2 = st.columns([3.2, 1.2])
        with c_t1:
            st.markdown(
                '<div style="margin-bottom: 14px;">'
                '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">📦 Dashboard Principal de Estoque</h1>'
                '<p style="color:#687787; font-size:0.9rem; margin:0;">Monitoramento consolidado de insumos, capital imobilizado e níveis de reposição.</p>'
                "</div>",
                unsafe_allow_html=True,
            )
        with c_t2:
            st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
            pdf_est_dash = gerar_pdf_relatorio_resumido(
                titulo="Dashboard Consolidado de Estoque",
                periodo_str=datetime.now().strftime("%d/%m/%Y"),
                kpis=[
                    ("Contagens Ativas (Saldo > 0)", str(itens_contados), f"Catálogo total: {len(df_estoque)} insumos"),
                    ("Itens em Alerta / Falta", str(len(df_crit)), "Abaixo da faixa mínima de segurança"),
                    ("Itens com Estoque OK", str(itens_ok), "Acima do ponto de reposição"),
                    ("Capital Imobilizado", formatar_moeda(capital_estoque), "Soma Qtd Atual × Preço Unitário"),
                ],
                df_resumo=df_estoque[df_estoque["Qtd_Atual"] > 0][
                    ["ID", "Item", "Categoria", "Unidade", "Qtd_Atual", "Preco_Unitario", "Status"]
                ]
                if itens_contados > 0
                else None,
            )
            st.download_button(
                label="📑 Exportar Relatório Resumido (PDF)",
                data=pdf_est_dash,
                file_name=f"resumo_estoque_mazon_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="btn_pdf_dash_estoque",
            )

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Contagens Ativas</div>'
                f'<div class="mazon-kpi-value">{itens_contados}</div>'
                f'<div class="mazon-kpi-sub">Itens com saldo físico &gt; 0</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Alertas de Reposição</div>'
                f'<div class="mazon-kpi-value" style="color:#D9383A;">{len(df_crit)}</div>'
                '<div class="mazon-kpi-sub">Crítico / Falta ou Alerta</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with k3:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Itens Abastecidos (OK)</div>'
                f'<div class="mazon-kpi-value" style="color:#1F9D55;">{itens_ok}</div>'
                '<div class="mazon-kpi-sub">Dentro da meta operacional</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Capital em Estoque</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(capital_estoque)}</div>'
                '<div class="mazon-kpi-sub">Valor financeiro total em estoque</div>'
                "</div>",
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)

        g1, g2 = st.columns(2)
        with g1:
            st.markdown(
                '<div class="mazon-card" style="margin-bottom:8px; padding:1.1rem 1.4rem 0.4rem 1.4rem;">'
                '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">Capital em Estoque por Categoria (R$)</h4>'
                "</div>",
                unsafe_allow_html=True,
            )
            if capital_estoque > 0:
                df_cap_cat = (
                    df_estoque.groupby("Categoria", as_index=False)["Valor_Total_Item"]
                    .sum()
                    .sort_values("Valor_Total_Item", ascending=True)
                )
                fig_cap = px.bar(
                    df_cap_cat,
                    x="Valor_Total_Item",
                    y="Categoria",
                    orientation="h",
                    color_discrete_sequence=["#0B1B2B"],
                    text=df_cap_cat["Valor_Total_Item"].apply(formatar_moeda),
                )
                fig_cap.update_traces(textposition="outside")
                fig_cap.update_layout(
                    paper_bgcolor="#FFFFFF",
                    plot_bgcolor="#FFFFFF",
                    font=dict(color="#0B1B2B", family="Poppins"),
                    xaxis=dict(title="Valor (R$)", gridcolor="#EBE4D8"),
                    yaxis=dict(title=""),
                    margin=dict(l=10, r=90, t=15, b=15),
                    height=330,
                )
                st.plotly_chart(fig_cap, use_container_width=True)
            else:
                st.info("Capital em estoque zerado (R$ 0,00). Cadastre custos e saldos para visualizar o gráfico.")

        with g2:
            st.markdown(
                '<div class="mazon-card" style="margin-bottom:8px; padding:1.1rem 1.4rem 0.4rem 1.4rem;">'
                '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">Quantidade Total Contada por Categoria</h4>'
                "</div>",
                unsafe_allow_html=True,
            )
            df_ativos = df_estoque[df_estoque["Qtd_Atual"] > 0]
            if not df_ativos.empty:
                df_qtd_cat = df_ativos.groupby("Categoria", as_index=False)["Qtd_Atual"].sum()
                fig_q = px.pie(
                    df_qtd_cat,
                    names="Categoria",
                    values="Qtd_Atual",
                    hole=0.58,
                    color_discrete_sequence=["#0B1B2B", "#D4A24C", "#1C354D", "#C59237", "#526375"],
                )
                fig_q.update_traces(textinfo="percent+label", textfont_color="#0B1B2B")
                fig_q.update_layout(
                    paper_bgcolor="#FFFFFF",
                    plot_bgcolor="#FFFFFF",
                    font=dict(color="#0B1B2B", family="Poppins"),
                    showlegend=False,
                    margin=dict(l=10, r=10, t=15, b=15),
                    height=330,
                )
                st.plotly_chart(fig_q, use_container_width=True)
            else:
                st.info("Nenhuma quantidade contada (0). Registre entradas ou contagens na tabela de estoque.")

    # --------------------------------------------------------------------------
    # SUB-PÁGINA 2: GESTÃO E TABELA DE ESTOQUE
    # --------------------------------------------------------------------------
    elif subpagina == "› Gestão e Tabela de Estoque":
        st.markdown(
            '<div style="margin-bottom: 14px;">'
            '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">📦 Gestão e Tabela Analítica de Estoque</h1>'
            '<p style="color:#687787; font-size:0.9rem; margin:0;">Controle completo de saldos, entradas/saídas, parâmetros de reposição e exportação multiformato.</p>'
            "</div>",
            unsafe_allow_html=True,
        )

        col_f1, col_f2, col_f3, col_btn = st.columns([2, 2, 1.8, 1.6])
        with col_f1:
            busca_texto = st.text_input("Buscar por Item ou ID", placeholder="Digite o nome ou código do insumo")
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
            if pode_editar:
                abrir_modal_cadastro = st.button("＋ Cadastrar / Editar Item", use_container_width=True)
            else:
                abrir_modal_cadastro = False
                st.markdown(
                    '<div style="font-size:0.76rem; color:#687787; padding-top:8px;">🔒 Edição restrita a Dev/Gerente</div>',
                    unsafe_allow_html=True,
                )

        @st.dialog("Cadastrar ou Editar Cadastro de Insumo")
        def modal_cadastro_estoque() -> None:
            st.markdown("Preencha os parâmetros do insumo para atualizar a aba `Estoque` da planilha unificada.")
            with st.form("form_modal_estoque_rbac", clear_on_submit=False):
                c1, c2 = st.columns(2)
                with c1:
                    in_id = st.text_input("ID do Insumo (opcional para novo)", placeholder="Ex: MZ-001")
                    in_nome = st.text_input("Nome do Item *", placeholder="Informe o nome do insumo")
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

        if abrir_modal_cadastro and pode_editar:
            modal_cadastro_estoque()

        with st.expander("🔄 Registrar Entrada ou Saída de Mercadoria (Movimentação Operacional)", expanded=False):
            with st.form("form_movimentacao_estoque"):
                cm1, cm2, cm3, cm4 = st.columns([2.2, 1.4, 1.4, 2])
                with cm1:
                    mov_item = st.selectbox("Insumo", options=df_estoque["Item"].tolist())
                with cm2:
                    mov_tipo = st.selectbox("Tipo de Movimento", options=["Entrada (+)", "Saída / Baixa (-)"])
                with cm3:
                    mov_qtd = st.number_input("Quantidade Movimentada", min_value=0.0, value=0.0, step=0.5)
                with cm4:
                    mov_obs = st.text_input("Motivo / NF / Observação", placeholder="Informe o motivo ou documento")

                if st.form_submit_button("Confirmar Movimentação de Estoque"):
                    if mov_qtd <= 0.0:
                        st.error("Informe uma quantidade maior que zero.")
                    else:
                        linha_item = df_estoque[df_estoque["Item"] == mov_item].iloc[0]
                        qtd_antiga = float(linha_item["Qtd_Atual"])
                        nova_qtd = (
                            qtd_antiga + float(mov_qtd)
                            if mov_tipo.startswith("Entrada")
                            else max(0.0, qtd_antiga - float(mov_qtd))
                        )
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
                                acao=f"{mov_tipo} de {mov_qtd} {linha_item['Unidade']} em '{mov_item}' ({qtd_antiga} -> {nova_qtd}) | Obs: {mov_obs or '-'}",
                                modulo="Estoque",
                            )
                            st.success(f"Saldo de '{mov_item}' atualizado para {nova_qtd:.2f} {linha_item['Unidade']}!")
                            st.rerun()

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

        # Botões de Exportação Multiformato (CSV, XLSX e PDF)
        df_export_est = df_filtrado[
            ["ID", "Item", "Categoria", "Unidade", "Qtd_Atual", "Est_Minimo", "Est_Maximo", "Preco_Unitario", "Fornecedor", "Status", "Data_Ultima_Reposicao"]
        ].copy()
        renderizar_barra_exportacao_tabela(
            df=df_export_est,
            prefixo_arquivo="estoque_mazon",
            titulo_pdf="Tabela Analítica de Estoque",
            subtitulo_pdf=f"Itens filtrados: {len(df_export_est)} | Capital filtrado: {formatar_moeda(float(df_filtrado['Valor_Total_Item'].sum()))}",
            key_suffix="tab_estoque_principal",
        )

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

        st.markdown(
            '<div class="mazon-table-wrapper">'
            '<table class="mazon-table">'
            "<thead><tr>"
            "<th>ID</th><th>Item</th><th>Categoria</th><th>Unidade</th>"
            "<th>Qtd Atual</th><th>Mínimo</th><th>Máximo</th><th>Custo Unit.</th>"
            "<th>Status</th><th>Última Reposição</th>"
            "</tr></thead>"
            f"<tbody>{corpo_tabela_html}</tbody>"
            "</table></div>",
            unsafe_allow_html=True,
        )

    # --------------------------------------------------------------------------
    # SUB-PÁGINA 3: COMPARATIVO DE PREÇOS DE FORNECEDORES (SEM DADOS FICTÍCIOS)
    # --------------------------------------------------------------------------
    elif subpagina == "› Comparativo de Preços":
        st.markdown(
            '<div style="margin-bottom: 14px;">'
            '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">⚖️ Comparativo de Preços entre Fornecedores</h1>'
            '<p style="color:#687787; font-size:0.9rem; margin:0;">Compare cotações reais de insumos para identificar o menor custo unitário e otimizar as compras.</p>'
            "</div>",
            unsafe_allow_html=True,
        )

        if "cotacoes_fornecedores" not in st.session_state:
            st.session_state["cotacoes_fornecedores"] = []

        with st.expander("➕ Registrar Cotação de Fornecedor para Comparação", expanded=True):
            with st.form("form_nova_cotacao", clear_on_submit=True):
                cc1, cc2, cc3, cc4 = st.columns([2, 1.5, 1.5, 1.5])
                with cc1:
                    cot_item = st.selectbox("Insumo", options=df_estoque["Item"].tolist())
                with cc2:
                    cot_forn1 = st.text_input("Fornecedor A", placeholder="Nome do Fornecedor A")
                    cot_preco1 = st.number_input("Preço Fornecedor A (R$)", min_value=0.0, value=0.0, step=0.5)
                with cc3:
                    cot_forn2 = st.text_input("Fornecedor B", placeholder="Nome do Fornecedor B")
                    cot_preco2 = st.number_input("Preço Fornecedor B (R$)", min_value=0.0, value=0.0, step=0.5)
                with cc4:
                    cot_forn3 = st.text_input("Fornecedor C (opcional)", placeholder="Nome do Fornecedor C")
                    cot_preco3 = st.number_input("Preço Fornecedor C (R$)", min_value=0.0, value=0.0, step=0.5)

                if st.form_submit_button("Adicionar ao Comparativo de Preços", use_container_width=True):
                    precos_validos = [
                        (cot_forn1.strip() or "Fornecedor A", cot_preco1)
                        for _ in [1]
                        if cot_preco1 > 0
                    ]
                    if cot_preco2 > 0:
                        precos_validos.append((cot_forn2.strip() or "Fornecedor B", cot_preco2))
                    if cot_preco3 > 0:
                        precos_validos.append((cot_forn3.strip() or "Fornecedor C", cot_preco3))

                    if not precos_validos:
                        st.error("Informe pelo menos um preço maior que R$ 0,00 para comparar.")
                    else:
                        melhor_forn, menor_p = min(precos_validos, key=lambda x: x[1])
                        maior_p = max(p for _, p in precos_validos)
                        economia = maior_p - menor_p
                        st.session_state["cotacoes_fornecedores"].append(
                            {
                                "Insumo": cot_item,
                                "Fornecedor_A": f"{cot_forn1 or '-'}: {formatar_moeda(cot_preco1)}",
                                "Fornecedor_B": f"{cot_forn2 or '-'}: {formatar_moeda(cot_preco2)}",
                                "Fornecedor_C": f"{cot_forn3 or '-'}: {formatar_moeda(cot_preco3)}",
                                "Melhor_Fornecedor": melhor_forn,
                                "Melhor_Preco": menor_p,
                                "Economia_Unitaria": economia,
                            }
                        )
                        registrar_log(
                            usuario=USUARIO_ATUAL,
                            perfil=PERFIL_ATUAL,
                            acao=f"Cotação comparativa registrada para '{cot_item}' (Melhor: {melhor_forn} por {formatar_moeda(menor_p)})",
                            modulo="Estoque",
                        )
                        st.rerun()

        df_cot = pd.DataFrame(st.session_state["cotacoes_fornecedores"])
        if not df_cot.empty:
            renderizar_barra_exportacao_tabela(
                df=df_cot,
                prefixo_arquivo="comparativo_precos_mazon",
                titulo_pdf="Comparativo de Preços entre Fornecedores",
                subtitulo_pdf=f"Total de insumos cotados: {len(df_cot)}",
                key_suffix="tab_comp_precos",
            )
            st.dataframe(df_cot, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma cotação registrada na sessão (R$ 0,00). Utilize o formulário acima para comparar fornecedores.")

    # --------------------------------------------------------------------------
    # SUB-PÁGINA 4: CENTRAL OCR / IA DE NOTAS E CUPONS (SEM DADOS FICTÍCIOS)
    # --------------------------------------------------------------------------
    elif subpagina == "› Central OCR / IA":
        st.markdown(
            '<div style="margin-bottom: 14px;">'
            '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">🤖 Central OCR / IA — Leitura de Notas &amp; Atualização de Estoque</h1>'
            '<p style="color:#687787; font-size:0.9rem; margin:0;">Envie uma Nota Fiscal/Cupom ou cole os itens da nota para atualizar o estoque e arquivar no Google Drive.</p>'
            "</div>",
            unsafe_allow_html=True,
        )

        col_ocr1, col_ocr2 = st.columns([1.2, 1.2])
        with col_ocr1:
            st.markdown('<div class="mazon-card">', unsafe_allow_html=True)
            st.markdown("#### 📄 1. Anexar Documento Fiscal para Arquivamento no Drive")
            arq_ocr = st.file_uploader(
                "Selecione a Nota Fiscal ou Cupom (.pdf, .png, .jpg)",
                type=["pdf", "png", "jpg", "jpeg"],
                key="up_ocr_ia",
            )
            forn_ocr = st.text_input("Fornecedor da Nota", placeholder="Informe o fornecedor")
            texto_ocr = st.text_area(
                "Texto dos itens (formato: Insumo ; Quantidade ; Preço Unitário)",
                placeholder="Exemplo de linha real:\nSalmão ; 10 ; 68.50",
                height=130,
            )
            st.markdown("</div>", unsafe_allow_html=True)

        with col_ocr2:
            st.markdown('<div class="mazon-card">', unsafe_allow_html=True)
            st.markdown("#### ⚡ 2. Processamento e Sincronização com a Planilha")
            linhas_validas = []
            if texto_ocr.strip():
                for linha in texto_ocr.strip().splitlines():
                    partes = [p.strip() for p in re.split(r"[;|,\t]", linha) if p.strip()]
                    if len(partes) >= 2:
                        try:
                            q_val = float(partes[1].replace(",", "."))
                            p_val = float(partes[2].replace(",", ".")) if len(partes) >= 3 else 0.0
                            linhas_validas.append({"Item": partes[0], "Quantidade_Entrada": q_val, "Preco_Unitario": p_val})
                        except ValueError:
                            pass

            df_preview_ocr = pd.DataFrame(linhas_validas)
            if not df_preview_ocr.empty:
                st.dataframe(df_preview_ocr, use_container_width=True, hide_index=True)
                if st.button("Confirmar Entrada no Estoque & Arquivar Nota", use_container_width=True):
                    if arq_ocr is not None:
                        db.upload_arquivo_drive(
                            arquivo_bytes=arq_ocr.read(),
                            nome_arquivo=arq_ocr.name,
                            mime_type=arq_ocr.type or "application/octet-stream",
                            subpasta="Notas_Fiscais",
                        )
                    atualizados = 0
                    for _, r_ocr in df_preview_ocr.iterrows():
                        masc = df_estoque["Item"].astype(str).str.lower() == str(r_ocr["Item"]).lower()
                        if masc.any():
                            item_row = df_estoque[masc].iloc[0]
                            nova_q = float(item_row["Qtd_Atual"]) + float(r_ocr["Quantidade_Entrada"])
                            novo_p = float(r_ocr["Preco_Unitario"]) if float(r_ocr["Preco_Unitario"]) > 0 else float(item_row["Preco_Unitario"])
                            db.cadastrar_ou_atualizar_item_estoque(
                                item_id=str(item_row["ID"]),
                                nome_item=str(item_row["Item"]),
                                categoria=str(item_row["Categoria"]),
                                unidade=str(item_row["Unidade"]),
                                qtd_atual=nova_q,
                                est_minimo=float(item_row["Est_Minimo"]),
                                est_maximo=float(item_row["Est_Maximo"]),
                                preco_referencia=novo_p,
                                fornecedor_referencia=forn_ocr.strip() or str(item_row["Fornecedor"]),
                            )
                            atualizados += 1
                    registrar_log(
                        usuario=USUARIO_ATUAL,
                        perfil=PERFIL_ATUAL,
                        acao=f"Processamento OCR/IA: {atualizados} insumos atualizados no estoque (Fornecedor: {forn_ocr or '-'})",
                        modulo="Estoque",
                    )
                    st.success(f"{atualizados} insumos sincronizados com sucesso na aba Estoque!")
                    st.rerun()
            else:
                st.info("Nenhum item em fila de processamento (0 itens). Envie um documento ou informe os itens ao lado.")
            st.markdown("</div>", unsafe_allow_html=True)

    # --------------------------------------------------------------------------
    # SUB-PÁGINA 5: GERADOR DE LISTA DE COMPRAS (EXPORTAÇÃO CSV, XLSX E PDF)
    # --------------------------------------------------------------------------
    elif subpagina == "› Gerador de Lista de Compras":
        st.markdown(
            '<div style="margin-bottom: 14px;">'
            '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">🛒 Gerador Oficial de Lista de Compras</h1>'
            '<p style="color:#687787; font-size:0.9rem; margin:0;">Cálculo automático de reposição para insumos abaixo do estoque mínimo com exportação em CSV, XLSX e PDF.</p>'
            "</div>",
            unsafe_allow_html=True,
        )

        df_alerta = df_estoque[df_estoque["Status"].isin(["CRÍTICO/FALTA", "ALERTA"])].copy()
        df_alerta["Qtd_a_Comprar"] = (df_alerta["Est_Maximo"] - df_alerta["Qtd_Atual"]).clip(lower=0.0)
        df_alerta["Custo_Estimado"] = df_alerta["Qtd_a_Comprar"] * df_alerta["Preco_Unitario"]
        orcamento_total = float(df_alerta["Custo_Estimado"].sum()) if not df_alerta.empty else 0.0

        ck1, ck2, ck3 = st.columns(3)
        with ck1:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Itens para Reposição</div>'
                f'<div class="mazon-kpi-value">{len(df_alerta)}</div>'
                '<div class="mazon-kpi-sub">Insumos em status Crítico ou Alerta</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with ck2:
            qtd_tot_comprar = float(df_alerta["Qtd_a_Comprar"].sum()) if not df_alerta.empty else 0.0
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Volume Total a Comprar</div>'
                f'<div class="mazon-kpi-value">{qtd_tot_comprar:.1f}</div>'
                '<div class="mazon-kpi-sub">Soma das unidades necessárias</div>'
                "</div>",
                unsafe_allow_html=True,
            )
        with ck3:
            st.markdown(
                '<div class="mazon-kpi-card">'
                '<div class="mazon-kpi-label">Orçamento Estimado</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(orcamento_total)}</div>'
                '<div class="mazon-kpi-sub">Baseado no último custo unitário</div>'
                "</div>",
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        df_export_compras = (
            df_alerta[["ID", "Item", "Categoria", "Unidade", "Qtd_Atual", "Est_Minimo", "Est_Maximo", "Qtd_a_Comprar", "Preco_Unitario", "Custo_Estimado", "Fornecedor"]].copy()
            if not df_alerta.empty
            else pd.DataFrame(columns=["ID", "Item", "Categoria", "Unidade", "Qtd_Atual", "Est_Minimo", "Est_Maximo", "Qtd_a_Comprar", "Preco_Unitario", "Custo_Estimado", "Fornecedor"])
        )

        renderizar_barra_exportacao_tabela(
            df=df_export_compras,
            prefixo_arquivo="lista_compras_mazon",
            titulo_pdf="Lista Oficial de Compras & Reposição",
            subtitulo_pdf=f"Total de Itens: {len(df_export_compras)} | Orçamento Estimado: {formatar_moeda(orcamento_total)}",
            key_suffix="tab_lista_compras",
        )

        if not df_export_compras.empty:
            st.dataframe(df_export_compras, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum insumo em alerta de compra no momento (0 itens | Orçamento: R$ 0,00).")


# ==============================================================================
# MÓDULO 3: TAB 💰 GESTÃO FINANCEIRA (6 SUB-PÁGINAS + RESTRIÇÃO CAIXA + FILTROS)
# ==============================================================================
def renderizar_modulo_financeiro(subpagina: str) -> None:
    """
    Renderiza as 6 sub-páginas financeiras:
    1. Lançamentos e NFs (Upload/Visualização)
    2. Faturamento
    3. DRE (Demonstrativo do Resultado)
    4. CMV (Custo de Mercadoria Vendida)
    5. Break-Even (Ponto de Equilíbrio)
    6. Relatório Descritivo
    - Perfil Caixa é restrito exclusivamente à sub-página 1 ("Lançamentos e NFs").
    - Trava Estrita de Exclusão: Visível e operável EXCLUSIVAMENTE para st.session_state["usuario"] == "Arthur".
    """
    # Bloqueio rigoroso para perfil Caixa caso tente acessar outra sub-página
    if PERFIL_ATUAL == "Caixa" and not tem_permissao("fin_relatorios"):
        subpagina = "› Lançamentos e NFs"

    # Visão Restrita para Caixa na sub-página Lançamentos e NFs (Sem gráficos, DRE ou margens)
    if PERFIL_ATUAL == "Caixa" and not tem_permissao("fin_relatorios"):
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
                cx_valor = st.number_input("Valor Bruto Apurado (R$) *", min_value=0.0, value=0.0, step=10.0)
                cx_desc = st.text_input("Observação / Turno", placeholder="Informe o turno ou observação")
                cx_arq = st.file_uploader(
                    "Anexar Comprovante de Fechamento (opcional - PDF/JPG/PNG)",
                    type=["pdf", "png", "jpg", "jpeg"],
                    key="up_caixa_fechamento",
                )

                if st.form_submit_button("Registrar Faturamento no Sistema", use_container_width=True):
                    if cx_valor <= 0.0:
                        st.error("Informe um valor bruto maior que R$ 0,00.")
                    else:
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
                nf_subpasta = st.selectbox("Pasta de Destino no Google Drive", options=["Notas_Fiscais", "Comprovantes"])
                nf_forn = st.text_input("Fornecedor / Emissor *", placeholder="Informe o fornecedor")
                nf_cat = st.selectbox(
                    "Categoria da Despesa / Nota",
                    options=[c for c in CATEGORIAS_FINANCEIRAS if not c.startswith("Faturamento")],
                )
                nf_forma = st.selectbox("Forma de Pagamento", options=FORMAS_PAGAMENTO, key="fp_nf_cx")
                nf_valor = st.number_input("Valor Total do Documento (R$) *", min_value=0.0, value=0.0, step=10.0)
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
                    elif nf_valor <= 0.0:
                        st.error("Informe um valor maior que R$ 0,00.")
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
    # PAINEL DE FILTROS INTERATIVOS NO TOPO (PARA DEV / GERENTE EM TODAS AS SUB-PÁGINAS)
    # ==========================================================================
    df_fin = db.ler_aba("Financeiro").copy()
    df_fin["Valor"] = pd.to_numeric(df_fin["Valor"], errors="coerce").fillna(0.0)
    df_fin["Data_dt"] = pd.to_datetime(df_fin["Data"], errors="coerce").dt.date

    st.markdown('<div class="mazon-card" style="padding: 1.05rem 1.35rem; margin-bottom: 1rem;">', unsafe_allow_html=True)
    st.markdown(
        '<div style="font-size:0.76rem; text-transform:uppercase; letter-spacing:1px; color:#9A6B1F; font-weight:600; margin-bottom:8px;">🔎 Painel de Filtros Financeiros Interativos (Tempo Real)</div>',
        unsafe_allow_html=True,
    )

    datas_validas = df_fin["Data_dt"].dropna()
    dt_min_padrao = datas_validas.min() if not datas_validas.empty else date.today()
    dt_max_padrao = datas_validas.max() if not datas_validas.empty else date.today()
    if dt_min_padrao > dt_max_padrao:
        dt_min_padrao = dt_max_padrao

    fornecedores_opcoes = sorted(df_fin["Fornecedor"].astype(str).unique().tolist()) if not df_fin.empty else []
    categorias_opcoes = sorted(df_fin["Categoria"].astype(str).unique().tolist()) if not df_fin.empty else CATEGORIAS_FINANCEIRAS
    formas_opcoes = sorted(df_fin["Forma_Pagamento"].astype(str).unique().tolist()) if not df_fin.empty else FORMAS_PAGAMENTO
    val_max_global = float(df_fin["Valor"].max()) if not df_fin.empty else 0.0

    f_col1, f_col2, f_col3, f_col4 = st.columns([1.4, 1.4, 1.5, 1.5])
    with f_col1:
        periodo_sel = st.date_input("Período (Inicial — Final)", value=(dt_min_padrao, dt_max_padrao))
    with f_col2:
        fornecedores_sel = st.multiselect("Fornecedor / Origem", options=fornecedores_opcoes, default=[], placeholder="Todos")
    with f_col3:
        categorias_sel = st.multiselect("Categoria Financeira", options=categorias_opcoes, default=[], placeholder="Todas")
    with f_col4:
        formas_sel = st.multiselect("Forma de Pagamento", options=formas_opcoes, default=[], placeholder="Todas")

    if val_max_global > 0.0:
        faixa_valor = st.slider(
            "Faixa de Valor do Lançamento (R$)",
            min_value=0.0,
            max_value=float(val_max_global),
            value=(0.0, float(val_max_global)),
            step=10.0,
        )
    else:
        st.caption("Faixa de Valor do Lançamento: R$ 0,00 — R$ 0,00 (sem lançamentos registrados)")
        faixa_valor = (0.0, 0.0)
    st.markdown("</div>", unsafe_allow_html=True)

    # Aplicação dos filtros
    df_filtrado = df_fin.copy()
    periodo_texto = f"{dt_min_padrao.strftime('%d/%m/%Y')} a {dt_max_padrao.strftime('%d/%m/%Y')}"
    if isinstance(periodo_sel, (list, tuple)) and len(periodo_sel) == 2:
        dt_ini, dt_fim = periodo_sel[0], periodo_sel[1]
        periodo_texto = f"{dt_ini.strftime('%d/%m/%Y')} a {dt_fim.strftime('%d/%m/%Y')}"
        df_filtrado = df_filtrado[(df_filtrado["Data_dt"] >= dt_ini) & (df_filtrado["Data_dt"] <= dt_fim)]
    elif isinstance(periodo_sel, date):
        periodo_texto = f"A partir de {periodo_sel.strftime('%d/%m/%Y')}"
        df_filtrado = df_filtrado[df_filtrado["Data_dt"] >= periodo_sel]

    if fornecedores_sel:
        df_filtrado = df_filtrado[df_filtrado["Fornecedor"].isin(fornecedores_sel)]
    if categorias_sel:
        df_filtrado = df_filtrado[df_filtrado["Categoria"].isin(categorias_sel)]
    if formas_sel:
        df_filtrado = df_filtrado[df_filtrado["Forma_Pagamento"].isin(formas_sel)]
    if val_max_global > 0.0:
        df_filtrado = df_filtrado[(df_filtrado["Valor"] >= faixa_valor[0]) & (df_filtrado["Valor"] <= faixa_valor[1])]

    # Métricas consolidadas do período filtrado
    df_receitas = df_filtrado[df_filtrado["Tipo"] == "Receita"].copy()
    df_despesas = df_filtrado[df_filtrado["Tipo"] == "Despesa"].copy()

    faturamento_bruto = float(df_receitas["Valor"].sum()) if not df_receitas.empty else 0.0
    impostos_taxas = float(
        df_despesas[df_despesas["Categoria"] == "Impostos & Taxas de Cartão/Apps"]["Valor"].sum()
    ) if not df_despesas.empty else 0.0
    receita_liquida = faturamento_bruto - impostos_taxas

    categorias_cmv = [
        "Insumos - Proteínas e Pescados",
        "Insumos - Mercearia e Hortifruti",
        "Insumos - Embalagens e Descartáveis",
    ]
    cmv_insumos = float(
        df_despesas[
            (df_despesas["Categoria"].isin(categorias_cmv))
            | (df_despesas["Natureza_Custo"] == "Custo Variável")
        ]["Valor"].sum()
    ) if not df_despesas.empty else 0.0

    custos_fixos = float(
        df_despesas[df_despesas["Natureza_Custo"] == "Custo Fixo"]["Valor"].sum()
    ) if not df_despesas.empty else 0.0
    despesas_totais = float(df_despesas["Valor"].sum()) if not df_despesas.empty else 0.0
    lucro_liquido = faturamento_bruto - despesas_totais
    margem_liquida = (lucro_liquido / faturamento_bruto * 100.0) if faturamento_bruto > 0 else 0.0
    margem_contribuicao = receita_liquida - cmv_insumos
    indice_mc = (margem_contribuicao / faturamento_bruto * 100.0) if faturamento_bruto > 0 else 0.0
    cmv_percentual = (cmv_insumos / faturamento_bruto * 100.0) if faturamento_bruto > 0 else 0.0
    ponto_equilibrio = (custos_fixos / (indice_mc / 100.0)) if indice_mc > 0 else 0.0

    # --------------------------------------------------------------------------
    # SUB-PÁGINA 1: LANÇAMENTOS E NFs (UPLOAD / VISUALIZAÇÃO / EXCLUSÃO ARTHUR)
    # --------------------------------------------------------------------------
    if subpagina == "› Lançamentos e NFs":
        st.markdown(
            '<div style="margin-bottom: 14px;">'
            '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">💰 Lançamentos Financeiros &amp; Gestão de Notas Fiscais</h1>'
            '<p style="color:#687787; font-size:0.9rem; margin:0;">Registro de receitas, despesas, arquivamento em nuvem no Google Drive e exportação multiformato.</p>'
            "</div>",
            unsafe_allow_html=True,
        )

        with st.expander("➕ Novo Lançamento Financeiro / Upload de Nota Fiscal ou Comprovante (Google Drive)", expanded=False):
            with st.form("form_lancamento_completo"):
                c_l1, c_l2, c_l3, c_l4 = st.columns(4)
                with c_l1:
                    l_data = st.date_input("Data", value=date.today())
                    l_tipo = st.selectbox("Tipo *", options=["Receita", "Despesa"])
                with c_l2:
                    l_natureza = st.selectbox("Natureza", options=["Faturamento Bruto", "Custo Variável", "Custo Fixo"])
                    l_cat = st.selectbox("Categoria *", options=CATEGORIAS_FINANCEIRAS)
                with c_l3:
                    l_forn = st.text_input("Fornecedor / Origem", placeholder="Informe o fornecedor ou cliente")
                    l_forma = st.selectbox("Forma de Pagamento", options=FORMAS_PAGAMENTO)
                with c_l4:
                    l_valor = st.number_input("Valor (R$) *", min_value=0.0, value=0.0, step=10.0)
                    l_subpasta = st.selectbox("Subpasta no Drive", options=["Notas_Fiscais", "Comprovantes"])

                l_desc = st.text_input("Descrição Detalhada *", placeholder="Informe a descrição do lançamento")
                l_arq = st.file_uploader(
                    "Anexar Nota Fiscal ou Comprovante (.pdf, .png, .jpg) — Upload para a pasta principal no Google Drive",
                    type=["pdf", "png", "jpg", "jpeg"],
                )

                if st.form_submit_button("Salvar Lançamento & Arquivar Documento", use_container_width=True):
                    if l_valor <= 0.0:
                        st.error("Informe um valor maior que R$ 0,00.")
                    else:
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

        df_exp_fin = df_filtrado.drop(columns=["Data_dt"], errors="ignore").copy()
        renderizar_barra_exportacao_tabela(
            df=df_exp_fin,
            prefixo_arquivo="lancamentos_financeiros_mazon",
            titulo_pdf="Relatório Analítico de Lançamentos Financeiros & NFs",
            subtitulo_pdf=f"Período: {periodo_texto} | Receitas: {formatar_moeda(faturamento_bruto)} | Despesas: {formatar_moeda(despesas_totais)}",
            key_suffix="tab_fin_lancamentos",
        )

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
            '<tr><td colspan="11" style="text-align:center; color:#687787; padding:18px;">Nenhum lançamento encontrado (R$ 0,00).</td></tr>'
        )

        st.markdown(
            '<div class="mazon-table-wrapper">'
            '<table class="mazon-table">'
            "<thead><tr>"
            "<th>ID</th><th>Data</th><th>Tipo</th><th>Natureza</th><th>Descrição</th>"
            "<th>Categoria</th><th>Fornecedor</th><th>Pagamento</th><th>Valor</th>"
            "<th>Comprovante / NF</th><th>Responsável</th>"
            "</tr></thead>"
            f"<tbody>{corpo_fin}</tbody>"
            "</table></div>",
            unsafe_allow_html=True,
        )

        # TRAVA ESTRITA DE EXCLUSÃO: EXCLUSIVAMENTE QUANDO st.session_state["usuario"] == "Arthur"
        if st.session_state.get("usuario") == "Arthur":
            with st.expander("🗑️ Exclusão de Lançamentos Financeiros / Notas Fiscais (Exclusivo Arthur - Dev)", expanded=False):
                st.warning(
                    "Trava Anti-Fraude Ativa: Apenas o usuário Arthur (Dev) pode visualizar e executar a exclusão definitiva de lançamentos."
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
                else:
                    st.info("Não há lançamentos cadastrados para excluir.")

    # --------------------------------------------------------------------------
    # SUB-PÁGINA 2: FATURAMENTO
    # --------------------------------------------------------------------------
    elif subpagina == "› Faturamento":
        fat_salao = float(df_receitas[df_receitas["Categoria"] == "Faturamento Salão"]["Valor"].sum()) if not df_receitas.empty else 0.0
        fat_delivery = float(df_receitas[df_receitas["Categoria"].astype(str).str.contains("Delivery")]["Valor"].sum()) if not df_receitas.empty else 0.0
        ticket_medio_lanc = (faturamento_bruto / len(df_receitas)) if not df_receitas.empty else 0.0

        cf1, cf2 = st.columns([3.2, 1.2])
        with cf1:
            st.markdown(
                '<div style="margin-bottom: 14px;">'
                '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">📈 Análise de Faturamento &amp; Canais de Venda</h1>'
                f'<p style="color:#687787; font-size:0.9rem; margin:0;">Desempenho de receitas brutas no período selecionado ({periodo_texto}).</p>'
                "</div>",
                unsafe_allow_html=True,
            )
        with cf2:
            st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
            pdf_fat = gerar_pdf_relatorio_resumido(
                titulo="Relatório Resumido de Faturamento",
                periodo_str=periodo_texto,
                kpis=[
                    ("Faturamento Bruto Total", formatar_moeda(faturamento_bruto), f"{len(df_receitas)} lançamentos de receita"),
                    ("Faturamento Salão", formatar_moeda(fat_salao), "Receita presencial no restaurante"),
                    ("Faturamento Delivery", formatar_moeda(fat_delivery), "Vendas via iFood / Delivery próprio"),
                    ("Média por Fechamento", formatar_moeda(ticket_medio_lanc), "Faturamento médio por lançamento"),
                ],
                df_resumo=df_receitas[["ID_Lancamento", "Data", "Categoria", "Forma_Pagamento", "Valor", "Responsavel"]] if not df_receitas.empty else None,
            )
            st.download_button(
                label="📑 Exportar Relatório Resumido (PDF)",
                data=pdf_fat,
                file_name=f"relatorio_faturamento_mazon_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="btn_pdf_faturamento",
            )

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Faturamento Bruto Total</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(faturamento_bruto)}</div>'
                f'<div class="mazon-kpi-sub">{len(df_receitas)} registros no período</div></div>',
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Faturamento Salão</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(fat_salao)}</div>'
                '<div class="mazon-kpi-sub">Atendimento presencial</div></div>',
                unsafe_allow_html=True,
            )
        with k3:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Faturamento Delivery</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(fat_delivery)}</div>'
                '<div class="mazon-kpi-sub">iFood &amp; Canais Digitais</div></div>',
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Média por Lançamento</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(ticket_medio_lanc)}</div>'
                '<div class="mazon-kpi-sub">Média aritmética dos fechamentos</div></div>',
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
        g1, g2 = st.columns(2)
        with g1:
            st.markdown(
                '<div class="mazon-card" style="margin-bottom:8px; padding:1.1rem 1.4rem 0.4rem 1.4rem;">'
                '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">Evolução Diária do Faturamento (R$)</h4></div>',
                unsafe_allow_html=True,
            )
            if not df_receitas.empty and faturamento_bruto > 0:
                df_fat_dia = df_receitas.groupby("Data", as_index=False)["Valor"].sum().sort_values("Data")
                fig_f_dia = px.bar(df_fat_dia, x="Data", y="Valor", color_discrete_sequence=["#0B1B2B"])
                fig_f_dia.update_layout(
                    paper_bgcolor="#FFFFFF",
                    plot_bgcolor="#FFFFFF",
                    font=dict(color="#0B1B2B", family="Poppins"),
                    height=320,
                    margin=dict(l=12, r=12, t=15, b=15),
                )
                st.plotly_chart(fig_f_dia, use_container_width=True)
            else:
                st.info("Faturamento zerado no período selecionado (R$ 0,00).")

        with g2:
            st.markdown(
                '<div class="mazon-card" style="margin-bottom:8px; padding:1.1rem 1.4rem 0.4rem 1.4rem;">'
                '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">Faturamento por Forma de Pagamento</h4></div>',
                unsafe_allow_html=True,
            )
            if not df_receitas.empty and faturamento_bruto > 0:
                df_fp = df_receitas.groupby("Forma_Pagamento", as_index=False)["Valor"].sum()
                fig_fp = px.pie(
                    df_fp,
                    names="Forma_Pagamento",
                    values="Valor",
                    hole=0.58,
                    color_discrete_sequence=["#0B1B2B", "#D4A24C", "#1C354D", "#C59237", "#526375"],
                )
                fig_fp.update_traces(textinfo="percent+label", textfont_color="#0B1B2B")
                fig_fp.update_layout(
                    paper_bgcolor="#FFFFFF",
                    plot_bgcolor="#FFFFFF",
                    font=dict(color="#0B1B2B", family="Poppins"),
                    showlegend=False,
                    height=320,
                    margin=dict(l=10, r=10, t=15, b=15),
                )
                st.plotly_chart(fig_fp, use_container_width=True)
            else:
                st.info("Sem receitas por forma de pagamento no período (R$ 0,00).")

    # --------------------------------------------------------------------------
    # SUB-PÁGINA 3: DRE (DEMONSTRATIVO DO RESULTADO DO EXERCÍCIO)
    # --------------------------------------------------------------------------
    elif subpagina == "› DRE (Demonstrativo do Resultado)":
        df_dre_tabela = pd.DataFrame(
            [
                {"Linha DRE": "1. (+) Receita Operacional Bruta", "Valor (R$)": faturamento_bruto, "Análise Vertical (%)": 100.0 if faturamento_bruto > 0 else 0.0},
                {"Linha DRE": "2. (-) Impostos e Taxas de Cartão/Apps", "Valor (R$)": -impostos_taxas, "Análise Vertical (%)": (impostos_taxas / faturamento_bruto * 100.0) if faturamento_bruto > 0 else 0.0},
                {"Linha DRE": "3. (=) Receita Operacional Líquida", "Valor (R$)": receita_liquida, "Análise Vertical (%)": (receita_liquida / faturamento_bruto * 100.0) if faturamento_bruto > 0 else 0.0},
                {"Linha DRE": "4. (-) CMV / Custos Variáveis de Insumos", "Valor (R$)": -cmv_insumos, "Análise Vertical (%)": cmv_percentual},
                {"Linha DRE": "5. (=) Margem de Contribuição", "Valor (R$)": margem_contribuicao, "Análise Vertical (%)": indice_mc},
                {"Linha DRE": "6. (-) Custos e Despesas Fixas Operacionais", "Valor (R$)": -custos_fixos, "Análise Vertical (%)": (custos_fixos / faturamento_bruto * 100.0) if faturamento_bruto > 0 else 0.0},
                {"Linha DRE": "7. (=) Resultado Líquido do Período (Lucro / Prejuízo)", "Valor (R$)": lucro_liquido, "Análise Vertical (%)": margem_liquida},
            ]
        )

        cd1, cd2 = st.columns([3.2, 1.2])
        with cd1:
            st.markdown(
                '<div style="margin-bottom: 14px;">'
                '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">📊 DRE — Demonstrativo do Resultado do Exercício</h1>'
                f'<p style="color:#687787; font-size:0.9rem; margin:0;">Apuração contábil-gerencial consolidada ({periodo_texto}).</p>'
                "</div>",
                unsafe_allow_html=True,
            )
        with cd2:
            st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
            pdf_dre = gerar_pdf_relatorio_resumido(
                titulo="DRE — Demonstrativo do Resultado",
                periodo_str=periodo_texto,
                kpis=[
                    ("Receita Bruta", formatar_moeda(faturamento_bruto), "100% da base de faturamento"),
                    ("CMV / Custos Variáveis", formatar_moeda(cmv_insumos), f"{cmv_percentual:.1f}% sobre o faturamento"),
                    ("Custos Fixos", formatar_moeda(custos_fixos), "Folha, aluguel, energia e operacional"),
                    ("Resultado Líquido", formatar_moeda(lucro_liquido), f"Margem Líquida: {margem_liquida:.1f}%"),
                ],
                df_resumo=df_dre_tabela,
            )
            st.download_button(
                label="📑 Exportar Relatório Resumido (PDF)",
                data=pdf_dre,
                file_name=f"dre_mazon_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="btn_pdf_dre_resumo",
            )

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Receita Bruta</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(faturamento_bruto)}</div>'
                f'<div class="mazon-kpi-sub">Receita Líquida: {formatar_moeda(receita_liquida)}</div></div>',
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Margem de Contribuição</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(margem_contribuicao)}</div>'
                f'<div class="mazon-kpi-sub">Índice MC: {indice_mc:.1f}%</div></div>',
                unsafe_allow_html=True,
            )
        with k3:
            cor_l = "#1F9D55" if lucro_liquido >= 0 else "#D9383A"
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Resultado Líquido</div>'
                f'<div class="mazon-kpi-value" style="color:{cor_l};">{formatar_moeda(lucro_liquido)}</div>'
                f'<div class="mazon-kpi-sub">Despesas Totais: {formatar_moeda(despesas_totais)}</div></div>',
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Margem Líquida (%)</div>'
                f'<div class="mazon-kpi-value" style="color:{cor_l};">{margem_liquida:.1f}%</div>'
                '<div class="mazon-kpi-sub">Lucro Líquido / Receita Bruta</div></div>',
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)

        col_dre_g, col_dre_t = st.columns([1.2, 1.2])
        with col_dre_g:
            fig_dre = go.Figure(
                go.Waterfall(
                    name="DRE",
                    orientation="v",
                    measure=["relative", "relative", "relative", "relative", "total"],
                    x=["Receita Bruta", "(-) Impostos/Taxas", "(-) CMV/Variáveis", "(-) Custos Fixos", "(=) Lucro Líquido"],
                    textposition="outside",
                    text=[
                        formatar_moeda(faturamento_bruto),
                        f"-{formatar_moeda(impostos_taxas)}",
                        f"-{formatar_moeda(cmv_insumos)}",
                        f"-{formatar_moeda(custos_fixos)}",
                        formatar_moeda(lucro_liquido),
                    ],
                    y=[faturamento_bruto, -impostos_taxas, -cmv_insumos, -custos_fixos, lucro_liquido],
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
                height=340,
                margin=dict(l=15, r=15, t=25, b=15),
            )
            st.plotly_chart(fig_dre, use_container_width=True)

        with col_dre_t:
            df_dre_vis = df_dre_tabela.copy()
            df_dre_vis["Valor (R$)"] = df_dre_vis["Valor (R$)"].apply(formatar_moeda)
            df_dre_vis["Análise Vertical (%)"] = df_dre_vis["Análise Vertical (%)"].apply(lambda v: f"{v:.1f}%")
            st.dataframe(df_dre_vis, use_container_width=True, hide_index=True, height=340)

    # --------------------------------------------------------------------------
    # SUB-PÁGINA 4: CMV (CUSTO DE MERCADORIA VENDIDA)
    # --------------------------------------------------------------------------
    elif subpagina == "› CMV (Custo de Mercadoria Vendida)":
        df_estoque_cmv = db.obter_estoque_enriquecido()
        saldo_estoque_atual = float(df_estoque_cmv["Valor_Total_Item"].sum()) if not df_estoque_cmv.empty else 0.0

        ccmv1, ccmv2 = st.columns([3.2, 1.2])
        with ccmv1:
            st.markdown(
                '<div style="margin-bottom: 14px;">'
                '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">🍣 CMV — Custo de Mercadoria Vendida</h1>'
                f'<p style="color:#687787; font-size:0.9rem; margin:0;">Análise de compras de insumos e impacto percentual sobre a receita ({periodo_texto}).</p>'
                "</div>",
                unsafe_allow_html=True,
            )
        with ccmv2:
            st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
            pdf_cmv = gerar_pdf_relatorio_resumido(
                titulo="Relatório de CMV (Custo de Mercadoria Vendida)",
                periodo_str=periodo_texto,
                kpis=[
                    ("Compras de Insumos / CMV (R$)", formatar_moeda(cmv_insumos), "Total de custos variáveis de mercadorias"),
                    ("Índice CMV (%)", f"{cmv_percentual:.1f}%", "CMV / Faturamento Bruto"),
                    ("Faturamento Base", formatar_moeda(faturamento_bruto), "Receita bruta no período filtrado"),
                    ("Saldo Atual em Estoque", formatar_moeda(saldo_estoque_atual), "Valor atual inventariado na aba Estoque"),
                ],
            )
            st.download_button(
                label="📑 Exportar Relatório Resumido (PDF)",
                data=pdf_cmv,
                file_name=f"relatorio_cmv_mazon_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="btn_pdf_cmv_resumo",
            )

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">CMV Apurado no Período</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(cmv_insumos)}</div>'
                '<div class="mazon-kpi-sub">Insumos e custos variáveis</div></div>',
                unsafe_allow_html=True,
            )
        with k2:
            cor_cmv = "#1F9D55" if cmv_percentual <= 35.0 else "#D9383A"
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">CMV sobre Faturamento (%)</div>'
                f'<div class="mazon-kpi-value" style="color:{cor_cmv};">{cmv_percentual:.1f}%</div>'
                '<div class="mazon-kpi-sub">Referência ideal gastronômica: 28% a 35%</div></div>',
                unsafe_allow_html=True,
            )
        with k3:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Faturamento Bruto</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(faturamento_bruto)}</div>'
                '<div class="mazon-kpi-sub">Base de cálculo do período</div></div>',
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Inventário Atual em Estoque</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(saldo_estoque_atual)}</div>'
                '<div class="mazon-kpi-sub">Capital físico avaliado na aba Estoque</div></div>',
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
        df_cmv_cat = df_despesas[
            (df_despesas["Categoria"].isin(categorias_cmv))
            | (df_despesas["Natureza_Custo"] == "Custo Variável")
        ]
        if not df_cmv_cat.empty and cmv_insumos > 0:
            res_cmv = df_cmv_cat.groupby("Categoria", as_index=False)["Valor"].sum()
            fig_cmv = px.bar(
                res_cmv,
                x="Categoria",
                y="Valor",
                color_discrete_sequence=["#0B1B2B"],
                text=res_cmv["Valor"].apply(formatar_moeda),
            )
            fig_cmv.update_layout(
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
                font=dict(color="#0B1B2B", family="Poppins"),
                height=330,
            )
            st.plotly_chart(fig_cmv, use_container_width=True)
        else:
            st.info("Nenhum custo de mercadoria/insumo registrado no período selecionado (R$ 0,00 | CMV: 0.0%).")

    # --------------------------------------------------------------------------
    # SUB-PÁGINA 5: BREAK-EVEN (PONTO DE EQUILÍBRIO)
    # --------------------------------------------------------------------------
    elif subpagina == "› Break-Even (Ponto de Equilíbrio)":
        margem_seguranca = (
            ((faturamento_bruto - ponto_equilibrio) / faturamento_bruto * 100.0)
            if (faturamento_bruto > 0 and ponto_equilibrio > 0)
            else 0.0
        )

        cbe1, cbe2 = st.columns([3.2, 1.2])
        with cbe1:
            st.markdown(
                '<div style="margin-bottom: 14px;">'
                '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">🎯 Break-Even Point (Ponto de Equilíbrio Operacional)</h1>'
                f'<p style="color:#687787; font-size:0.9rem; margin:0;">Faturamento mínimo necessário para cobrir 100% dos custos fixos e variáveis ({periodo_texto}).</p>'
                "</div>",
                unsafe_allow_html=True,
            )
        with cbe2:
            st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
            pdf_be = gerar_pdf_relatorio_resumido(
                titulo="Análise de Break-Even (Ponto de Equilíbrio)",
                periodo_str=periodo_texto,
                kpis=[
                    ("Ponto de Equilíbrio (R$)", formatar_moeda(ponto_equilibrio), "Meta mínima de faturamento para lucro zero"),
                    ("Faturamento Realizado", formatar_moeda(faturamento_bruto), f"Margem de Segurança: {margem_seguranca:.1f}%"),
                    ("Custos Fixos Totais", formatar_moeda(custos_fixos), "Despesas estruturais do período"),
                    ("Índice Margem de Contribuição", f"{indice_mc:.1f}%", "Percentual livre para cobrir custos fixos"),
                ],
            )
            st.download_button(
                label="📑 Exportar Relatório Resumido (PDF)",
                data=pdf_be,
                file_name=f"relatorio_breakeven_mazon_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="btn_pdf_be_resumo",
            )

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Ponto de Equilíbrio (Break-Even)</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(ponto_equilibrio)}</div>'
                '<div class="mazon-kpi-sub">Custos Fixos / Índice MC</div></div>',
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Faturamento Atual</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(faturamento_bruto)}</div>'
                f'<div class="mazon-kpi-sub">Diferença: {formatar_moeda(faturamento_bruto - ponto_equilibrio)}</div></div>',
                unsafe_allow_html=True,
            )
        with k3:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Custos Fixos Apurados</div>'
                f'<div class="mazon-kpi-value">{formatar_moeda(custos_fixos)}</div>'
                '<div class="mazon-kpi-sub">Base fixa a ser coberta</div></div>',
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                '<div class="mazon-kpi-card"><div class="mazon-kpi-label">Margem de Segurança (%)</div>'
                f'<div class="mazon-kpi-value">{margem_seguranca:.1f}%</div>'
                '<div class="mazon-kpi-sub">Acima/abaixo do ponto de equilíbrio</div></div>',
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
        df_be_comp = pd.DataFrame(
            [
                {"Métrica": "Faturamento Realizado", "Valor": faturamento_bruto},
                {"Métrica": "Ponto de Equilíbrio (Meta Mínima)", "Valor": ponto_equilibrio},
                {"Métrica": "Custos Fixos", "Valor": custos_fixos},
                {"Métrica": "Custos Variáveis / CMV", "Valor": cmv_insumos},
            ]
        )
        fig_be = px.bar(
            df_be_comp,
            x="Métrica",
            y="Valor",
            color="Métrica",
            color_discrete_sequence=["#0B1B2B", "#D4A24C", "#526375", "#9A6B1F"],
            text=df_be_comp["Valor"].apply(formatar_moeda),
        )
        fig_be.update_layout(
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            font=dict(color="#0B1B2B", family="Poppins"),
            showlegend=False,
            height=330,
        )
        st.plotly_chart(fig_be, use_container_width=True)

    # --------------------------------------------------------------------------
    # SUB-PÁGINA 6: RELATÓRIO DESCRITIVO
    # --------------------------------------------------------------------------
    elif subpagina == "› Relatório Descritivo":
        qtd_com_anexo = int(df_filtrado["Link_Comprovante"].astype(str).str.startswith("http").sum()) if not df_filtrado.empty else 0
        secoes_desc = [
            (
                "2. Síntese de Faturamento e Receitas",
                f"No período de {periodo_texto}, o Restaurante Mazon registrou {len(df_receitas)} lançamentos de receita, "
                f"totalizando um Faturamento Bruto de {formatar_moeda(faturamento_bruto)} e Receita Líquida de {formatar_moeda(receita_liquida)}.",
            ),
            (
                "3. Estrutura de Custos, CMV e Despesas Operacionais",
                f"As despesas somaram {formatar_moeda(despesas_totais)} ({len(df_despesas)} registros), sendo {formatar_moeda(cmv_insumos)} "
                f"referentes a CMV/Custos Variáveis ({cmv_percentual:.1f}% da receita) e {formatar_moeda(custos_fixos)} de Custos Fixos.",
            ),
            (
                "4. Resultado Líquido, Break-Even e Rastreabilidade de Comprovantes",
                f"O resultado líquido apurado foi de {formatar_moeda(lucro_liquido)} (Margem Líquida de {margem_liquida:.1f}%), "
                f"com Ponto de Equilíbrio estimado em {formatar_moeda(ponto_equilibrio)}. "
                f"Foram identificados {qtd_com_anexo} documentos com link direto no Google Drive.",
            ),
        ]

        crd1, crd2 = st.columns([3.2, 1.2])
        with crd1:
            st.markdown(
                '<div style="margin-bottom: 14px;">'
                '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">📑 Relatório Descritivo Executivo</h1>'
                f'<p style="color:#687787; font-size:0.9rem; margin:0;">Diagnóstico textual e analítico consolidado do período ({periodo_texto}).</p>'
                "</div>",
                unsafe_allow_html=True,
            )
        with crd2:
            st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
            pdf_desc = gerar_pdf_relatorio_resumido(
                titulo="Relatório Descritivo Executivo",
                periodo_str=periodo_texto,
                kpis=[
                    ("Faturamento Bruto", formatar_moeda(faturamento_bruto), f"{len(df_receitas)} receitas lançadas"),
                    ("Despesas Totais", formatar_moeda(despesas_totais), f"{len(df_despesas)} despesas lançadas"),
                    ("Resultado Líquido", formatar_moeda(lucro_liquido), f"Margem: {margem_liquida:.1f}%"),
                    ("Comprovantes no Google Drive", str(qtd_com_anexo), "Arquivos vinculados aos lançamentos"),
                ],
                secoes_texto=secoes_desc,
                df_resumo=df_filtrado.drop(columns=["Data_dt"], errors="ignore") if not df_filtrado.empty else None,
            )
            st.download_button(
                label="📑 Exportar Relatório Resumido (PDF)",
                data=pdf_desc,
                file_name=f"relatorio_descritivo_mazon_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="btn_pdf_rel_descritivo",
            )

        for titulo_s, texto_s in secoes_desc:
            st.markdown(
                f'<div class="mazon-card" style="margin-bottom: 12px;">'
                f'<h4 style="margin:0 0 6px 0; color:#0B1B2B;">{titulo_s}</h4>'
                f'<p style="margin:0; color:#526375; font-size:0.92rem; line-height:1.65;">{texto_s}</p>'
                "</div>",
                unsafe_allow_html=True,
            )


# ==============================================================================
# MÓDULO 4: TAB 🛡️ GESTÃO DE USUÁRIOS & AUDITORIA (HIERARQUIA + PERMISSÕES)
# ==============================================================================
def renderizar_modulo_auditoria_usuarios(subpagina: str) -> None:
    """
    Renderiza as 2 sub-páginas da Tab Gestão de Usuários & Auditoria:
    1. Log de Auditoria (Dev vê 100%, Gerente vê Estoque e Caixa; exportação CSV/XLSX/PDF)
    2. Gestão de Usuários & Permissões (Respeito estrito à Hierarquia Dev > Gerente + Matriz de Checkboxes)
    """
    # --------------------------------------------------------------------------
    # SUB-PÁGINA 1: LOG DE AUDITORIA
    # --------------------------------------------------------------------------
    if subpagina == "› Log de Auditoria":
        st.markdown(
            '<div style="margin-bottom: 14px;">'
            '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">🛡️ Logs de Auditoria &amp; Rastreabilidade</h1>'
            '<p style="color:#687787; font-size:0.9rem; margin:0;">Histórico de operações gravado na aba <code>Logs_Auditoria</code> com exportação em CSV, XLSX e PDF.</p>'
            "</div>",
            unsafe_allow_html=True,
        )

        df_logs = db.ler_aba("Logs_Auditoria").copy()

        if PERFIL_ATUAL != "Dev":
            st.info(
                "🔎 Filtro Hierárquico Ativo: Exibindo automaticamente apenas os eventos referentes aos módulos de **Estoque** e **Caixa**."
            )
            df_logs = df_logs[df_logs["Modulo"].astype(str).isin(["Estoque", "Caixa"])].copy()
        else:
            st.markdown(
                '<div class="mazon-card" style="padding:0.9rem 1.3rem;">'
                '<span class="badge-best-price">AUTORIDADE TOTAL DEV (ARTHUR)</span>'
                '<span style="margin-left:10px; font-size:0.86rem; color:#0B1B2B;">Visualizando 100% do Log Geral do Sistema (Autenticação, Usuários, Estoque, Caixa e Financeiro).</span>'
                "</div>",
                unsafe_allow_html=True,
            )

        col_a1, col_a2 = st.columns(2)
        with col_a1:
            busca_log = st.text_input("Filtrar por Usuário ou Ação", placeholder="Digite termo da ação ou usuário")
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

        df_logs = df_logs.iloc[::-1].copy()

        renderizar_barra_exportacao_tabela(
            df=df_logs,
            prefixo_arquivo="logs_auditoria_mazon",
            titulo_pdf="Relatório Oficial de Logs de Auditoria",
            subtitulo_pdf=f"Total de eventos auditados: {len(df_logs)} | Perfil solicitante: {PERFIL_ATUAL}",
            key_suffix="tab_logs_auditoria",
        )

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
            "<thead><tr>"
            "<th>Data / Hora</th><th>Usuário</th><th>Perfil</th><th>Ação Executada</th><th>Módulo</th>"
            "</tr></thead>"
            f"<tbody>{corpo_logs}</tbody>"
            "</table></div>",
            unsafe_allow_html=True,
        )

    # --------------------------------------------------------------------------
    # SUB-PÁGINA 2: GESTÃO DE USUÁRIOS, HIERARQUIA E MATRIZ DE PERMISSÕES
    # --------------------------------------------------------------------------
    elif subpagina == "› Gestão de Usuários & Permissões":
        st.markdown(
            '<div style="margin-bottom: 14px;">'
            '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.7rem;">🔐 Gestão de Usuários, Hierarquia &amp; Permissões Granulares</h1>'
            '<p style="color:#687787; font-size:0.9rem; margin:0;">Gerencie contas ativas/inativas, atribua cargos respeitando a hierarquia e personalize permissões por módulo.</p>'
            "</div>",
            unsafe_allow_html=True,
        )

        # Lista de usuários respeitando a Hierarquia (Gerente NÃO visualiza contas Dev)
        df_visiveis = db.listar_usuarios_visiveis(PERFIL_ATUAL)

        if PERFIL_ATUAL == "Dev":
            cargos_permitidos = ["Dev", "Gerente", "Caixa", "Equipe", "Personalizado"]
            st.markdown(
                '<div class="mazon-card" style="padding:0.85rem 1.25rem; margin-bottom:14px;">'
                '<span class="badge-best-price">HIERARQUIA DEV (ARTHUR)</span>'
                '<span style="margin-left:10px; font-size:0.86rem; color:#0B1B2B;">Autoridade máxima ativa: Você pode visualizar, criar e promover contas para qualquer nível, incluindo <strong>Dev</strong>.</span>'
                "</div>",
                unsafe_allow_html=True,
            )
        else:
            cargos_permitidos = ["Gerente", "Caixa", "Equipe", "Personalizado"]
            st.markdown(
                '<div class="mazon-card" style="padding:0.85rem 1.25rem; margin-bottom:14px;">'
                '<span class="badge-alerta">HIERARQUIA GERENTE</span>'
                '<span style="margin-left:10px; font-size:0.86rem; color:#0B1B2B;">Você pode criar e gerenciar perfis até o nível <strong>Gerente</strong>. Contas de nível <strong>Dev</strong> são protegidas e ocultas.</span>'
                "</div>",
                unsafe_allow_html=True,
            )

        # Tabela de Usuários Cadastrados (Ocultando senhas em texto aberto na tabela visual)
        st.markdown("#### 👥 Contas de Usuários Visíveis na Sua Hierarquia")
        linhas_u_html = []
        for _, r_u in df_visiveis.iterrows():
            st_ativo = str(r_u.get("Ativo", "Sim")).strip()
            badge_u = '<span class="badge-ok">Ativo</span>' if st_ativo.lower() == "sim" else '<span class="badge-critico">Inativo</span>'
            perms_codigos = db.resolver_permissoes(str(r_u["Perfil"]), str(r_u.get("Permissoes", "")))
            perms_legiveis = ", ".join(perms_codigos)
            linhas_u_html.append(
                "<tr>"
                f'<td style="font-weight:600; color:#0B1B2B;">{r_u["Usuario"]}</td>'
                f'<td><span class="badge-alerta">{r_u["Perfil"]}</span></td>'
                f"<td>{badge_u}</td>"
                f'<td>{r_u.get("Data_Cadastro", "-")}</td>'
                f'<td style="font-size:0.8rem; color:#526375;">{perms_legiveis}</td>'
                "</tr>"
            )

        st.markdown(
            '<div class="mazon-table-wrapper">'
            '<table class="mazon-table">'
            "<thead><tr>"
            "<th>Usuário</th><th>Cargo / Perfil</th><th>Status</th><th>Data Cadastro</th><th>Módulos Permitidos</th>"
            "</tr></thead>"
            f"<tbody>{''.join(linhas_u_html)}</tbody>"
            "</table></div>",
            unsafe_allow_html=True,
        )

        # Seleção de Usuário Existente ou Criação de Novo Usuário + Matriz de Seletores (Checkboxes)
        st.markdown('<div class="mazon-card">', unsafe_allow_html=True)
        st.markdown("#### ⚙️ Criar Novo Usuário ou Editar Conta / Permissões Existentes")

        opcoes_selecao = ["➕ Cadastrar Novo Usuário"] + df_visiveis["Usuario"].astype(str).tolist()
        modo_selecionado = st.selectbox(
            "Selecione um usuário existente para editar ou crie um novo:",
            options=opcoes_selecao,
            key="sel_usuario_edicao_rbac",
        )

        if modo_selecionado == "➕ Cadastrar Novo Usuário":
            u_nome_default = ""
            u_perfil_default = "Equipe"
            u_ativo_default = "Sim"
            perms_atuais_sel = list(PERMISSOES_POR_PERFIL["Equipe"])
        else:
            linha_u = df_visiveis[df_visiveis["Usuario"].astype(str) == modo_selecionado].iloc[0]
            u_nome_default = str(linha_u["Usuario"])
            u_perfil_default = str(linha_u["Perfil"]) if str(linha_u["Perfil"]) in cargos_permitidos else cargos_permitidos[0]
            u_ativo_default = str(linha_u.get("Ativo", "Sim"))
            perms_atuais_sel = db.resolver_permissoes(str(linha_u["Perfil"]), str(linha_u.get("Permissoes", "")))

        with st.form("form_rbac_matriz_permissoes", clear_on_submit=False):
            cu1, cu2, cu3, cu4 = st.columns(4)
            with cu1:
                in_u_nome = st.text_input(
                    "Nome de Usuário *",
                    value=u_nome_default,
                    placeholder="Digite o identificador do usuário",
                )
            with cu2:
                idx_cargo = cargos_permitidos.index(u_perfil_default) if u_perfil_default in cargos_permitidos else 0
                in_u_perfil = st.selectbox("Cargo / Nível Hierárquico *", options=cargos_permitidos, index=idx_cargo)
            with cu3:
                in_u_senha = st.text_input(
                    "Senha (deixe em branco para manter a atual ao editar)",
                    type="password",
                    placeholder="Digite a senha",
                )
            with cu4:
                idx_at = 0 if u_ativo_default.lower() == "sim" else 1
                in_u_ativo = st.selectbox("Status da Conta", options=["Sim", "Não"], index=idx_at)

            st.markdown(
                '<div style="font-size:0.8rem; font-weight:600; color:#9A6B1F; text-transform:uppercase; letter-spacing:0.8px; margin: 12px 0 6px 0;">🎛️ Matriz de Seletores de Permissões Customizadas por Módulo</div>',
                unsafe_allow_html=True,
            )
            st.caption(
                "Marque ou desmarque os módulos abaixo para personalizar o acesso desta conta (ou mantenha a seleção para aplicar as permissões selecionadas):"
            )

            col_chk1, col_chk2 = st.columns(2)
            chaves_perm = list(PERMISSOES_DISPONIVEIS.keys())
            permissoes_escolhidas: List[str] = []

            for i, chave_p in enumerate(chaves_perm):
                label_p = PERMISSOES_DISPONIVEIS[chave_p]
                col_alvo = col_chk1 if i % 2 == 0 else col_chk2
                with col_alvo:
                    marcado = st.checkbox(
                        label_p,
                        value=(chave_p in perms_atuais_sel),
                        key=f"chk_perm_{modo_selecionado}_{chave_p}",
                    )
                    if marcado:
                        permissoes_escolhidas.append(chave_p)

            if st.form_submit_button("💾 Salvar Usuário, Cargo e Matriz de Permissões", use_container_width=True):
                ok_save, msg_save = db.cadastrar_ou_atualizar_usuario(
                    novo_usuario=in_u_nome,
                    perfil=in_u_perfil,
                    senha=in_u_senha,
                    ativo=in_u_ativo,
                    permissoes_lista=permissoes_escolhidas,
                    operador_perfil=PERFIL_ATUAL,
                )
                if ok_save:
                    registrar_log(
                        usuario=USUARIO_ATUAL,
                        perfil=PERFIL_ATUAL,
                        acao=f"Gestão de acesso: usuário '{in_u_nome}' salvo com perfil '{in_u_perfil}', Ativo={in_u_ativo}, Permissões=[{','.join(permissoes_escolhidas)}]",
                        modulo="Usuarios",
                    )
                    # Se o usuário editou a própria conta, sincroniza na sessão atual
                    if in_u_nome.strip().lower() == USUARIO_ATUAL.lower():
                        st.session_state["permissoes"] = permissoes_escolhidas
                    st.success(msg_save)
                    st.rerun()
                else:
                    st.error(msg_save)
        st.markdown("</div>", unsafe_allow_html=True)


# ==============================================================================
# ROTEADOR PRINCIPAL EM CASCATA (ABA SUPERIOR ATIVA -> MÓDULO + SUB-PÁGINA)
# ==============================================================================
if ABA_SUPERIOR_ATIVA == "🏠 Home":
    renderizar_modulo_home()
elif ABA_SUPERIOR_ATIVA == "📦 Estoque":
    renderizar_modulo_estoque(SUBPAGINA_ESTOQUE)
elif ABA_SUPERIOR_ATIVA == "💰 Gestão Financeira":
    renderizar_modulo_financeiro(SUBPAGINA_FINANCEIRO)
elif ABA_SUPERIOR_ATIVA == "🛡️ Gestão de Usuários & Auditoria":
    renderizar_modulo_auditoria_usuarios(SUBPAGINA_AUDITORIA)
