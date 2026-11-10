"""
RESTAURANTE MAZON - Sistema Automatizado de Gestão de Estoque e Inteligência Financeira
FASE 1: Arquitetura em Nuvem (Google Sheets & Drive), Identidade Luxury Minimalist UI (Poppins)
e 5 Módulos Operacionais.
"""

from __future__ import annotations

import io
import os
from datetime import datetime
from typing import Optional

import pandas as pd
import plotly.express as px
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
    CATEGORIAS_OFICIAIS,
    MazonSheetsManager,
    processar_documento_ocr_ia,
)


def resolver_caminho_recurso(nome_arquivo: str) -> str:
    """Retorna o caminho válido do arquivo no diretório atual ou no pacote."""
    if os.path.exists(nome_arquivo):
        return nome_arquivo
    caminho_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), nome_arquivo)
    return caminho_script


# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA E INJEÇÃO DE CSS (LUXURY MINIMALIST UI)
# ==============================================================================
_caminho_icone = resolver_caminho_recurso("logo2.png")
if not os.path.exists(_caminho_icone):
    for _cand in ["logo2.jpg", "logo2.jpeg", "logo.png", "logo.jpeg", "logo.jpg"]:
        _teste = resolver_caminho_recurso(_cand)
        if os.path.exists(_teste):
            _caminho_icone = _teste
            break

_icone_aba = Image.open(_caminho_icone) if os.path.exists(_caminho_icone) else "🍣"

st.set_page_config(
    page_title="Restaurante Mazon | Gestão de Estoque & Inteligência Financeira",
    page_icon=_icone_aba,
    layout="wide",
    initial_sidebar_state="expanded",
)

# Força a limpeza do cache de memória na primeira execução da sessão
if "init_estoque_reset" not in st.session_state:
    st.session_state.estoque_df = None
    st.session_state["init_estoque_reset"] = True


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
    Injeta a fonte Google Fonts Poppins restrita APENAS a elementos de texto explícitos,
    protege integralmente os ícones Material do Streamlit e aplica o style.css.
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
        "div[data-testid='stSidebar'] div[role='radiogroup'] label > div:first-child, "
        "section[data-testid='stSidebar'] div[role='radiogroup'] label > div:first-child {"
        "display: none !important;"
        "}"
        "</style>"
    )
    st.markdown(regras_fonte_e_icones, unsafe_allow_html=True)

    caminho_css = resolver_caminho_recurso("style.css")
    if os.path.exists(caminho_css):
        with open(caminho_css, "r", encoding="utf-8") as f:
            css_content = f.read()
        st.markdown(f"<style>{css_content}</style>", unsafe_allow_html=True)


injetar_estilo_luxury_minimalist()

# Inicializa o Gerenciador de Banco de Dados Google Sheets / Drive
db = MazonSheetsManager()


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
    """
    Gera o relatório em PDF (via ReportLab) harmonizado na paleta Luxury Cream & Deep Navy:
    - Fundo de página suave (#FAF6F0) e faixa superior em Azul Marinho (#0B1B2B)
    - Imagem da logo oficial no topo via Image(caminho_logo, width=140, height=90)
    - Cabeçalhos de tabela em Azul Marinho (#0B1B2B) com texto Creme (#F4EDE3) e linhas em creme/dourado.
    """
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
        """Aplica fundo creme suave (#FAF6F0), borda superior marinho (#0B1B2B) e filete dourado (#D4A24C)."""
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
# MENU LATERAL (SIDEBAR MINIMALISTA #0B1B2B COM LOGO INTEGRADA E SETAS)
# ==============================================================================
caminho_logo = obter_caminho_logo()
if caminho_logo:
    st.sidebar.image(caminho_logo, use_container_width=True)
else:
    st.sidebar.markdown(
        '<div style="text-align: center; padding: 10px 0 16px 0;">'
        '<div style="font-family: \'Poppins\', sans-serif; font-size: 1.5rem; font-weight: 600; color: #F4EDE3; letter-spacing: 2px;">MAZON</div>'
        '<div style="font-size: 0.72rem; color: #D4A24C; text-transform: uppercase; letter-spacing: 1.4px; margin-top: 2px;">Restaurante Japonês</div>'
        "</div>",
        unsafe_allow_html=True,
    )

with st.sidebar:
    st.markdown(
        '<p style="font-size:0.7rem; text-transform:uppercase; letter-spacing:1.5px; color:#D4A24C; font-weight:500; margin: 6px 0 8px 16px; opacity:0.85;">Navegação</p>',
        unsafe_allow_html=True,
    )
    pagina_selecionada = st.radio(
        label="Módulos do Sistema",
        options=[
            "›  Dashboard Principal",
            "›  Gestão e Tabela de Estoque",
            "›  Comparativo de Preços",
            "›  Central OCR / IA",
            "›  Gerador de Lista de Compras",
        ],
        label_visibility="collapsed",
    )

    st.markdown('<hr style="border-color: rgba(255,255,255,0.07); margin: 18px 0;">', unsafe_allow_html=True)

    # Status Minimalista da Conexão Google Cloud
    cor_status_nuvem = "#1F9D55" if db.connected else "#D4A24C"
    titulo_nuvem = "Google Sheets Online" if db.connected else "Estoque Oficial Ativo"
    st.markdown(
        f'<div style="padding: 8px 14px; margin-bottom: 12px; border-left: 2px solid {cor_status_nuvem};">'
        f'<div style="font-size: 0.74rem; font-weight: 500; color: {cor_status_nuvem}; letter-spacing: 0.5px;">● {titulo_nuvem}</div>'
        f'<div style="font-size: 0.72rem; color: #C5D1DE; opacity: 0.75; margin-top: 3px;">{db.connection_message}</div>'
        "</div>",
        unsafe_allow_html=True,
    )

    st.markdown(
        '<p style="font-size:0.7rem; text-transform:uppercase; letter-spacing:1.4px; color:#D4A24C; font-weight:500; margin: 12px 0 6px 14px; opacity:0.85;">Ajuste Rápido</p>',
        unsafe_allow_html=True,
    )
    df_est_sidebar = db.ler_aba("Estoque")
    opcoes_itens = df_est_sidebar["Item"].tolist() if not df_est_sidebar.empty else []
    if opcoes_itens:
        item_rapido = st.selectbox("Insumo", options=opcoes_itens, key="sb_item_rapido")
        linha_atual = df_est_sidebar[df_est_sidebar["Item"] == item_rapido].iloc[0]
        nova_qtd_rapida = st.number_input(
            f"Quantidade ({linha_atual['Unidade']})",
            min_value=0.0,
            value=float(linha_atual["Qtd_Atual"]),
            step=0.5,
            key="sb_qtd_rapida",
        )
        if st.button("Sincronizar Quantidade", use_container_width=True):
            ok, msg = db.cadastrar_ou_atualizar_item_estoque(
                item_id=str(linha_atual["ID"]),
                nome_item=str(linha_atual["Item"]),
                categoria=str(linha_atual["Categoria"]),
                unidade=str(linha_atual["Unidade"]),
                qtd_atual=nova_qtd_rapida,
                est_minimo=float(linha_atual["Est_Minimo"]),
                est_maximo=float(linha_atual["Est_Maximo"]),
            )
            if ok:
                st.success("Estoque atualizado!")
                st.rerun()


# ==============================================================================
# PÁGINA 1: ›  DASHBOARD PRINCIPAL
# ==============================================================================
if pagina_selecionada == "›  Dashboard Principal":
    st.markdown(
        '<div style="margin-bottom: 20px;">'
        '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.75rem;">Painel Executivo &amp; Controle de Estoque</h1>'
        '<p style="color:#687787; font-size:0.9rem; margin:0;">Visão consolidada do inventário oficial e inteligência financeira do Restaurante Mazon</p>'
        "</div>",
        unsafe_allow_html=True,
    )

    df_estoque = db.obter_estoque_enriquecido()
    df_criticos = df_estoque[df_estoque["Status"].isin(["CRÍTICO/FALTA", "PENDENTE"])].copy()

    # 1. KPIs em Cards Brancos Minimalistas (Sem bordas grossas coloridas no topo)
    total_itens = len(df_estoque)
    total_criticos = len(df_criticos)
    valor_total_estoque = float(df_estoque["Valor_Total_Item"].sum())

    col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
    with col_kpi1:
        st.markdown(
            '<div class="mazon-kpi-card">'
            '<div class="mazon-kpi-label">Total de Insumos Oficiais</div>'
            f'<div class="mazon-kpi-value">{total_itens}</div>'
            f'<div class="mazon-kpi-sub">Distribuídos em {df_estoque["Categoria"].nunique()} categorias operacionais</div>'
            "</div>",
            unsafe_allow_html=True,
        )
    with col_kpi2:
        cor_kpi_crit = "#D9383A" if total_criticos > 0 else "#1F9D55"
        st.markdown(
            '<div class="mazon-kpi-card">'
            '<div class="mazon-kpi-label">Itens Pendentes / Críticos</div>'
            f'<div class="mazon-kpi-value" style="color: {cor_kpi_crit};">{total_criticos}</div>'
            f'<div class="mazon-kpi-sub">{(total_criticos / max(total_itens, 1) * 100):.1f}% do portfólio requer reposição ou contagem</div>'
            "</div>",
            unsafe_allow_html=True,
        )
    with col_kpi3:
        st.markdown(
            '<div class="mazon-kpi-card">'
            '<div class="mazon-kpi-label">Valor Total Acumulado em Estoque</div>'
            f'<div class="mazon-kpi-value">{formatar_moeda(valor_total_estoque)}</div>'
            '<div class="mazon-kpi-sub">Avaliado pela última cotação de fornecedor</div>'
            "</div>",
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)

    # 2. Gráficos em Cards Minimalistas Brancos
    base_cat = pd.DataFrame({"Categoria": CATEGORIAS_OFICIAIS})
    df_cat = base_cat.merge(
        df_estoque.groupby("Categoria", as_index=False)["Valor_Total_Item"].sum(),
        on="Categoria",
        how="left",
    ).fillna(0.0)
    df_cat = df_cat.sort_values("Valor_Total_Item", ascending=True)

    col_graf1, col_graf2 = st.columns([3, 2])

    with col_graf1:
        st.markdown(
            '<div class="mazon-card" style="margin-bottom: 8px; padding: 1.15rem 1.4rem 0.4rem 1.4rem;">'
            '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">Capital Investido por Categoria (R$)</h4>'
            '<p style="font-size:0.78rem; color:#687787; margin-top:2px;">Distribuição financeira segundo as cotações registradas</p>'
            "</div>",
            unsafe_allow_html=True,
        )
        fig_bar = px.bar(
            df_cat,
            x="Valor_Total_Item",
            y="Categoria",
            orientation="h",
            text=df_cat["Valor_Total_Item"].apply(formatar_moeda),
            color="Valor_Total_Item",
            color_continuous_scale=["#0B1B2B", "#1E3A56", "#D4A24C"],
        )
        fig_bar.update_traces(
            textposition="outside",
            marker_line_color="#0B1B2B",
            marker_line_width=0.4,
        )
        fig_bar.update_layout(
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            font=dict(color="#0B1B2B", family="Poppins"),
            xaxis=dict(title="Valor Investido (R$)", gridcolor="#EBE4D8"),
            yaxis=dict(title=""),
            coloraxis_showscale=False,
            margin=dict(l=12, r=105, t=15, b=20),
            height=390,
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    with col_graf2:
        st.markdown(
            '<div class="mazon-card" style="margin-bottom: 8px; padding: 1.15rem 1.4rem 0.4rem 1.4rem;">'
            '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">Distribuição de Insumos por Categoria</h4>'
            '<p style="font-size:0.78rem; color:#687787; margin-top:2px;">Proporção dos 90 itens oficiais no catálogo</p>'
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
            margin=dict(l=10, r=10, t=15, b=20),
            height=390,
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    # 3. Card Minimalista de Alertas Críticos / Itens Pendentes
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
            f'<div class="critical-hero-title">Insumos em Falta, Pendentes ou Abaixo do Mínimo ({len(df_criticos)} itens)</div>'
            '<span class="badge-best-price">Reposição / Contagem</span>'
            "</div>"
            '<p style="color:#687787; font-size:0.85rem; margin:0;">'
            "Os insumos abaixo aguardam contagem inicial ou atingiram nível crítico de ruptura operacional:"
            "</p>"
            f'<div class="critical-items-grid">{pills_html}</div>'
            "</div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="mazon-card">'
            '<h3 style="color:#1F9D55; margin:0; font-size:1.05rem;">Operação Estável — Nenhum Insumo Abaixo do Estoque Mínimo</h3>'
            '<p style="color:#687787; font-size:0.86rem; margin-top:4px; margin-bottom:0;">Todos os itens monitorados encontram-se acima do ponto de ruptura.</p>'
            "</div>",
            unsafe_allow_html=True,
        )


# ==============================================================================
# PÁGINA 2: ›  GESTÃO E TABELA DE ESTOQUE
# ==============================================================================
elif pagina_selecionada == "›  Gestão e Tabela de Estoque":
    st.markdown(
        '<div style="margin-bottom: 18px;">'
        '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.75rem;">Gestão e Tabela de Estoque</h1>'
        '<p style="color:#687787; font-size:0.9rem; margin:0;">Controle completo dos 90 insumos oficiais, níveis mínimo/máximo e sincronização direta com a aba <code>Estoque</code> do Google Sheets.</p>'
        "</div>",
        unsafe_allow_html=True,
    )

    df_estoque = db.obter_estoque_enriquecido()

    col_f1, col_f2, col_f3, col_btn = st.columns([2, 2, 2, 1.5])
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
        abrir_modal = st.button("＋ Novo / Editar Item", use_container_width=True)

    @st.dialog("Cadastrar ou Atualizar Insumo no Estoque")
    def modal_cadastro_estoque() -> None:
        st.markdown(
            "Preencha os dados abaixo. Se o **ID** ou **Nome do Item** já existir, o registro será atualizado."
        )
        with st.form("form_modal_estoque", clear_on_submit=False):
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
                in_preco = st.number_input(
                    "Preço Unitário de Referência (R$)", min_value=0.0, value=0.0, step=1.0
                )
            in_forn = st.text_input("Fornecedor de Referência", value="Fornecedor Padrão")

            submetido = st.form_submit_button("Salvar e Sincronizar com Google Sheets")
            if submetido:
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
                        preco_referencia=in_preco if in_preco > 0 else None,
                        fornecedor_referencia=in_forn,
                    )
                    if ok:
                        st.success(msg)
                        st.rerun()

    if abrir_modal:
        modal_cadastro_estoque()

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
            f"<td>{badge_html}</td>"
            f'<td style="color:#687787;">{row["Data_Ultima_Reposicao"]}</td>'
            "</tr>"
        )

    corpo_tabela_html = "".join(linhas_lista) if linhas_lista else (
        '<tr><td colspan="9" style="text-align:center; color:#687787; padding:18px;">Nenhum insumo encontrado para os filtros selecionados.</td></tr>'
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
        "<th>Status</th>"
        "<th>Última Reposição</th>"
        "</tr>"
        "</thead>"
        f"<tbody>{corpo_tabela_html}</tbody>"
        "</table>"
        "</div>"
    )
    st.markdown(tabela_estoque_html, unsafe_allow_html=True)

    with st.expander("Editar ou Cadastrar Insumo Diretamente na Página", expanded=False):
        with st.form("form_inline_estoque"):
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                f_id = st.text_input("ID (Ex: MZ-001)", value="")
                f_item = st.text_input("Item *", value="")
            with col2:
                f_cat = st.selectbox("Categoria", options=CATEGORIAS_OFICIAIS)
                f_un = st.text_input("Unidade", value="kg")
            with col3:
                f_qtd = st.number_input("Qtd Atual", min_value=0.0, value=0.0, step=0.5)
                f_min = st.number_input("Est. Mínimo", min_value=0.0, value=0.0, step=0.5)
            with col4:
                f_max = st.number_input("Est. Máximo", min_value=0.0, value=0.0, step=1.0)
                f_preco = st.number_input("Preço Unitário (R$)", min_value=0.0, value=0.0, step=1.0)

            if st.form_submit_button("Salvar Alterações no Google Sheets"):
                if f_item.strip():
                    ok, msg = db.cadastrar_ou_atualizar_item_estoque(
                        item_id=f_id,
                        nome_item=f_item,
                        categoria=f_cat,
                        unidade=f_un,
                        qtd_atual=f_qtd,
                        est_minimo=f_min,
                        est_maximo=f_max,
                        preco_referencia=f_preco if f_preco > 0 else None,
                    )
                    if ok:
                        st.success(msg)
                        st.rerun()
                else:
                    st.warning("Informe o nome do Item para salvar.")


# ==============================================================================
# PÁGINA 3: ›  COMPARATIVO DE PREÇOS
# ==============================================================================
elif pagina_selecionada == "›  Comparativo de Preços":
    st.markdown(
        '<div style="margin-bottom: 18px;">'
        '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.75rem;">Inteligência de Preços e Fornecedores</h1>'
        '<p style="color:#687787; font-size:0.9rem; margin:0;">Análise histórica de cotações, evolução de custo por insumo e recomendação automática do Melhor Preço Atual.</p>'
        "</div>",
        unsafe_allow_html=True,
    )

    df_estoque_lista = db.ler_aba("Estoque")
    todos_insumos_oficiais = sorted(df_estoque_lista["Item"].unique().tolist())

    df_precos = db.ler_aba("Fornecedores_Precos").copy()
    if not df_precos.empty:
        df_precos["Preco_Unitario"] = pd.to_numeric(
            df_precos["Preco_Unitario"], errors="coerce"
        ).fillna(0.0)

    with st.expander("Registrar Nova Cotação de Fornecedor", expanded=df_precos.empty):
        with st.form("form_nova_cotacao"):
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                cot_insumo = st.selectbox("Insumo", options=todos_insumos_oficiais)
            with c2:
                cot_forn = st.text_input("Fornecedor *", placeholder="Ex: Mar & Cia Imports")
            with c3:
                cot_preco = st.number_input("Preço Unitário (R$) *", min_value=0.01, value=10.0, step=0.5)
            with c4:
                cot_unid = st.text_input("Unidade", value="kg")

            if st.form_submit_button("Registrar Cotação"):
                if cot_forn.strip():
                    db.registrar_preco_fornecedor(
                        insumo=cot_insumo,
                        fornecedor=cot_forn,
                        preco_unitario=cot_preco,
                        unidade=cot_unid,
                    )
                    st.success("Cotação registrada com sucesso!")
                    st.rerun()
                else:
                    st.error("Informe o nome do fornecedor.")

    if df_precos.empty:
        st.markdown(
            '<div class="mazon-card">'
            '<h4 style="margin:0; color:#0B1B2B; font-size:0.98rem;">Nenhuma cotação registrada ainda</h4>'
            '<p style="color:#687787; font-size:0.86rem; margin-top:6px; margin-bottom:0;">Utilize o formulário acima para registrar os preços reais de compra e habilitar o comparativo automático entre fornecedores.</p>'
            "</div>",
            unsafe_allow_html=True,
        )
    else:
        insumos_cotados = sorted(df_precos["Insumo"].unique().tolist())
        insumo_padrao_idx = (
            insumos_cotados.index("Salmão") if "Salmão" in insumos_cotados else 0
        )

        col_sel, _ = st.columns([2, 1])
        with col_sel:
            insumo_selecionado = st.selectbox(
                "Selecione o Insumo para Análise de Evolução e Comparativo:",
                options=insumos_cotados,
                index=insumo_padrao_idx,
            )

        df_insumo = df_precos[df_precos["Insumo"] == insumo_selecionado].sort_values("Data_Registro")

        ultimas_por_fornecedor = (
            df_insumo.sort_values("Data_Registro", ascending=False)
            .groupby("Fornecedor", as_index=False)
            .first()
            .sort_values("Preco_Unitario", ascending=True)
        )

        if not ultimas_por_fornecedor.empty:
            melhor_linha = ultimas_por_fornecedor.iloc[0]
            pior_preco = float(ultimas_por_fornecedor["Preco_Unitario"].max())
            melhor_preco = float(melhor_linha["Preco_Unitario"])
            economia_unit = max(0.0, pior_preco - melhor_preco)

            st.markdown(
                '<div class="mazon-card" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:14px;">'
                "<div>"
                '<span class="badge-best-price">MELHOR PREÇO ATUAL</span>'
                f'<h3 style="margin:8px 0 4px 0; color:#0B1B2B; font-size:1.2rem;">{melhor_linha["Fornecedor"]} — <span style="color:#9A6B1F;">{formatar_moeda(melhor_preco)} / {melhor_linha["Unidade"]}</span></h3>'
                f'<div style="font-size:0.83rem; color:#687787;">Insumo: <strong>{insumo_selecionado}</strong> · Última cotação em {melhor_linha["Data_Registro"]}</div>'
                "</div>"
                '<div style="text-align:right;">'
                '<div style="font-size:0.75rem; color:#687787; text-transform:uppercase;">Economia vs. Maior Cotação</div>'
                f'<div style="font-size:1.35rem; font-weight:600; color:#1F9D55;">{formatar_moeda(economia_unit)} / {melhor_linha["Unidade"]}</div>'
                "</div>"
                "</div>",
                unsafe_allow_html=True,
            )

        st.markdown(f"#### Evolução Histórica de Preço — {insumo_selecionado}")
        fig_evolucao = px.line(
            df_insumo,
            x="Data_Registro",
            y="Preco_Unitario",
            color="Fornecedor",
            markers=True,
            color_discrete_sequence=["#0B1B2B", "#D4A24C", "#1F9D55", "#D9383A", "#687787"],
            labels={
                "Data_Registro": "Data da Cotação / Compra",
                "Preco_Unitario": "Preço Unitário (R$)",
                "Fornecedor": "Fornecedor",
            },
        )
        fig_evolucao.update_traces(line=dict(width=2.5), marker=dict(size=8))
        fig_evolucao.update_layout(
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            font=dict(color="#0B1B2B", family="Poppins"),
            xaxis=dict(gridcolor="#EBE4D8"),
            yaxis=dict(gridcolor="#EBE4D8", tickprefix="R$ "),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=20, r=20, t=40, b=20),
            height=380,
        )
        st.plotly_chart(fig_evolucao, use_container_width=True)

        st.markdown("#### Tabela Comparativa de Fornecedores")
        ultimas_geral = (
            df_precos.sort_values("Data_Registro", ascending=False)
            .groupby(["Insumo", "Fornecedor"], as_index=False)
            .first()
        )
        min_por_insumo = ultimas_geral.groupby("Insumo")["Preco_Unitario"].transform("min")
        ultimas_geral["Eh_Melhor"] = ultimas_geral["Preco_Unitario"] == min_por_insumo
        ultimas_geral = ultimas_geral.sort_values(["Insumo", "Preco_Unitario"])

        linhas_comp_list = []
        for _, row in ultimas_geral.iterrows():
            destaque_tag = (
                '<span class="badge-best-price">MELHOR PREÇO ATUAL</span>'
                if row["Eh_Melhor"]
                else '<span style="color:#687787; font-size:0.78rem;">Cotação Secundária</span>'
            )
            cor_preco = "#1F9D55" if row["Eh_Melhor"] else "#0B1B2B"
            linhas_comp_list.append(
                "<tr>"
                f'<td style="color:#9A6B1F; font-weight:600;">{row["ID"]}</td>'
                f'<td style="font-weight:500; color:#0B1B2B;">{row["Insumo"]}</td>'
                f'<td>{row["Fornecedor"]}</td>'
                f'<td style="font-weight:600; color:{cor_preco};">{formatar_moeda(float(row["Preco_Unitario"]))}</td>'
                f'<td>{row["Unidade"]}</td>'
                f'<td>{row["Data_Registro"]}</td>'
                f"<td>{destaque_tag}</td>"
                "</tr>"
            )

        tabela_comp_html = (
            '<div class="mazon-table-wrapper">'
            '<table class="mazon-table">'
            "<thead>"
            "<tr>"
            "<th>ID</th>"
            "<th>Insumo</th>"
            "<th>Fornecedor</th>"
            "<th>Preço Unitário</th>"
            "<th>Unidade</th>"
            "<th>Data Registro</th>"
            "<th>Recomendação de Compra</th>"
            "</tr>"
            "</thead>"
            f"<tbody>{''.join(linhas_comp_list)}</tbody>"
            "</table>"
            "</div>"
        )
        st.markdown(tabela_comp_html, unsafe_allow_html=True)


# ==============================================================================
# PÁGINA 4: ›  CENTRAL OCR / IA
# ==============================================================================
elif pagina_selecionada == "›  Central OCR / IA":
    st.markdown(
        '<div style="margin-bottom: 18px;">'
        '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.75rem;">Central de Ingestão Documental (OCR / IA)</h1>'
        '<p style="color:#687787; font-size:0.9rem; margin:0;">Envie fotos ou PDFs de Notas Fiscais de Compra e Extratos de Vendas do Dia. Os arquivos são arquivados no Google Drive e pré-processados para a Fase 2.</p>'
        "</div>",
        unsafe_allow_html=True,
    )

    col_up, col_info = st.columns([3, 2])

    with col_up:
        st.markdown("#### Upload de Documentos")
        tipo_doc = st.radio(
            "Selecione a Natureza do Documento:",
            options=["Nota Fiscal / Cupom de Compra", "Extrato de Vendas do Dia (PDV)"],
            horizontal=True,
        )

        arquivos_enviados = st.file_uploader(
            "Arraste e solte imagens (.png, .jpg, .jpeg) ou documentos (.pdf)",
            type=["png", "jpg", "jpeg", "pdf"],
            accept_multiple_files=True,
        )

        fornecedor_manual = ""
        valor_manual = 0.0
        if tipo_doc == "Nota Fiscal / Cupom de Compra":
            c_nf1, c_nf2 = st.columns(2)
            with c_nf1:
                fornecedor_manual = st.text_input(
                    "Fornecedor (opcional p/ indexação imediata)", placeholder="Ex: Mar & Cia Imports"
                )
            with c_nf2:
                valor_manual = st.number_input("Valor Total da Nota (R$)", min_value=0.0, value=0.0, step=10.0)

        if arquivos_enviados:
            if st.button("Arquivar no Google Drive & Executar Pré-Análise IA", use_container_width=True):
                for arq in arquivos_enviados:
                    bytes_arq = arq.read()
                    link_drive = db.upload_arquivo_drive(
                        arquivo_bytes=bytes_arq,
                        nome_arquivo=arq.name,
                        mime_type=arq.type or "application/octet-stream",
                    )
                    resultado_ocr = processar_documento_ocr_ia(
                        arquivo_bytes=bytes_arq,
                        nome_arquivo=arq.name,
                        tipo_documento=tipo_doc,
                    )

                    if tipo_doc == "Nota Fiscal / Cupom de Compra":
                        df_entradas = db.ler_aba("Entradas_Compras")
                        nova_entrada = pd.DataFrame(
                            [
                                {
                                    "ID_Nota": f"NF-{datetime.now().strftime('%Y%m%d-%H%M')}",
                                    "Data": datetime.now().strftime("%Y-%m-%d"),
                                    "Fornecedor": fornecedor_manual.strip() or "Pendente OCR Fase 2",
                                    "Valor_Total": valor_manual,
                                    "Link_Drive": link_drive,
                                }
                            ]
                        )
                        df_entradas = pd.concat([df_entradas, nova_entrada], ignore_index=True)
                        db.salvar_aba("Entradas_Compras", df_entradas)

                    st.success(f"Arquivo '{arq.name}' salvo! Link Drive: {link_drive}")
                    st.json(resultado_ocr)

    with col_info:
        st.markdown(
            '<div class="mazon-card">'
            '<h4 style="color:#0B1B2B; margin-top:0; font-size:0.98rem;">Arquitetura de Visão Computacional (Fase 2)</h4>'
            '<p style="font-size:0.85rem; color:#0B1B2B;">'
            "O módulo <code>processar_documento_ocr_ia()</code> em <code>sheets_manager.py</code> possui os contratos de dados prontos para:"
            "</p>"
            '<ul style="font-size:0.83rem; color:#687787; line-height:1.6;">'
            "<li><strong>Notas de Compra:</strong> Extração de CNPJ/Fornecedor, Itens, Quantidades e Preços Unitários, alimentando simultaneamente <code>Entradas_Compras</code>, <code>Estoque</code> e <code>Fornecedores_Precos</code>.</li>"
            "<li><strong>Extratos de Vendas:</strong> Leitura de pratos vendidos no dia e baixa automática de insumos cruzando com a aba <code>Ficha_Tecnica</code>.</li>"
            "</ul>"
            "</div>",
            unsafe_allow_html=True,
        )

    tab_ent, tab_sai, tab_ft = st.tabs(
        ["Histórico Entradas_Compras", "Histórico Saidas_Vendas", "Ficha_Tecnica Cadastrada"]
    )
    with tab_ent:
        st.dataframe(db.ler_aba("Entradas_Compras"), use_container_width=True, hide_index=True)
    with tab_sai:
        st.dataframe(db.ler_aba("Saidas_Vendas"), use_container_width=True, hide_index=True)
    with tab_ft:
        st.dataframe(db.ler_aba("Ficha_Tecnica"), use_container_width=True, hide_index=True)


# ==============================================================================
# PÁGINA 5: ›  GERADOR DE LISTA DE COMPRAS
# ==============================================================================
elif pagina_selecionada == "›  Gerador de Lista de Compras":
    st.markdown(
        '<div style="margin-bottom: 18px;">'
        '<h1 style="margin-bottom:2px; color:#0B1B2B; font-size:1.75rem;">Gerador Inteligente de Lista de Compras</h1>'
        '<p style="color:#687787; font-size:0.9rem; margin:0;">Cálculo automatizado de reposição pela fórmula <code>Qtd_Necessaria = Est_Maximo - Qtd_Atual</code> para todos os itens em estado Crítico, Pendente ou Alerta, com exportação direta em PDF e Planilha.</p>'
        "</div>",
        unsafe_allow_html=True,
    )

    df_estoque = db.obter_estoque_enriquecido()
    df_precos = db.ler_aba("Fornecedores_Precos")
    if not df_precos.empty:
        df_precos["Preco_Unitario"] = pd.to_numeric(
            df_precos["Preco_Unitario"], errors="coerce"
        ).fillna(0.0)

    incluir_todos_abaixo_max = st.checkbox(
        "Incluir também itens 'OK/CHEIO' que estejam abaixo do Estoque Máximo", value=False
    )

    if incluir_todos_abaixo_max:
        df_alerta = df_estoque[df_estoque["Qtd_Atual"] < df_estoque["Est_Maximo"]].copy()
    else:
        df_alerta = df_estoque[
            df_estoque["Status"].isin(["CRÍTICO/FALTA", "ALERTA", "PENDENTE"])
        ].copy()

    df_alerta["Qtd_Necessaria"] = (df_alerta["Est_Maximo"] - df_alerta["Qtd_Atual"]).clip(lower=0.0)

    if not df_precos.empty:
        melhores_forn = (
            df_precos.sort_values("Data_Registro", ascending=False)
            .groupby(["Insumo", "Fornecedor"], as_index=False)
            .first()
            .sort_values("Preco_Unitario", ascending=True)
            .groupby("Insumo", as_index=False)
            .first()
            .rename(
                columns={
                    "Insumo": "Item",
                    "Fornecedor": "Melhor_Fornecedor",
                    "Preco_Unitario": "Melhor_Preco_Unitario",
                }
            )
        )
        df_compras = df_alerta.merge(
            melhores_forn[["Item", "Melhor_Fornecedor", "Melhor_Preco_Unitario"]],
            on="Item",
            how="left",
        )
    else:
        df_compras = df_alerta.copy()
        df_compras["Melhor_Fornecedor"] = "A Cotar"
        df_compras["Melhor_Preco_Unitario"] = 0.0

    df_compras["Melhor_Fornecedor"] = df_compras["Melhor_Fornecedor"].fillna("A Cotar")
    df_compras["Melhor_Preco_Unitario"] = df_compras["Melhor_Preco_Unitario"].fillna(
        df_compras["Ultimo_Preco_Unitario"]
    )
    df_compras["Custo_Estimado_Reposicao"] = (
        df_compras["Qtd_Necessaria"] * df_compras["Melhor_Preco_Unitario"]
    )

    custo_total_compra = float(df_compras["Custo_Estimado_Reposicao"].sum())

    c_res1, c_res2, c_res3 = st.columns(3)
    with c_res1:
        st.markdown(
            '<div class="mazon-kpi-card">'
            '<div class="mazon-kpi-label">Itens na Lista de Compra</div>'
            f'<div class="mazon-kpi-value">{len(df_compras)}</div>'
            '<div class="mazon-kpi-sub">Insumos requerendo reposição ou parametrização</div>'
            "</div>",
            unsafe_allow_html=True,
        )
    with c_res2:
        st.markdown(
            '<div class="mazon-kpi-card">'
            '<div class="mazon-kpi-label">Orçamento Estimado de Compra</div>'
            f'<div class="mazon-kpi-value">{formatar_moeda(custo_total_compra)}</div>'
            '<div class="mazon-kpi-sub">Calculado pelo Melhor Preço Atual de cada item</div>'
            "</div>",
            unsafe_allow_html=True,
        )
    with c_res3:
        st.markdown(
            '<div class="mazon-kpi-card">'
            '<div class="mazon-kpi-label">Fornecedores Envolvidos</div>'
            f'<div class="mazon-kpi-value">{df_compras["Melhor_Fornecedor"].nunique()}</div>'
            '<div class="mazon-kpi-sub">Agrupados por melhor cotação ativa</div>'
            "</div>",
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)

    if df_compras.empty:
        st.info("Nenhum item precisa de reposição no momento.")
    else:
        linhas_compra_list = []
        for _, row in df_compras.iterrows():
            linhas_compra_list.append(
                "<tr>"
                f'<td style="color:#9A6B1F; font-weight:600;">{row["ID"]}</td>'
                f'<td style="font-weight:500; color:#0B1B2B;">{row["Item"]}</td>'
                f'<td>{row["Categoria"]}</td>'
                f'<td>{renderizar_badge_status(str(row["Status"]))}</td>'
                f'<td>{float(row["Qtd_Atual"]):.2f} {row["Unidade"]}</td>'
                f'<td>{float(row["Est_Minimo"]):.2f} {row["Unidade"]}</td>'
                f'<td>{float(row["Est_Maximo"]):.2f} {row["Unidade"]}</td>'
                f'<td style="font-weight:600; color:#9A6B1F;">+{float(row["Qtd_Necessaria"]):.2f} {row["Unidade"]}</td>'
                f'<td>{row["Melhor_Fornecedor"]}</td>'
                f'<td style="font-weight:600; color:#1F9D55;">{formatar_moeda(float(row["Custo_Estimado_Reposicao"]))}</td>'
                "</tr>"
            )

        tabela_compras_html = (
            '<div class="mazon-table-wrapper">'
            '<table class="mazon-table">'
            "<thead>"
            "<tr>"
            "<th>ID</th>"
            "<th>Insumo</th>"
            "<th>Categoria</th>"
            "<th>Status</th>"
            "<th>Qtd Atual</th>"
            "<th>Estoque Mínimo</th>"
            "<th>Estoque Máximo</th>"
            "<th>Qtd a Comprar</th>"
            "<th>Fornecedor</th>"
            "<th>Subtotal Est.</th>"
            "</tr>"
            "</thead>"
            f"<tbody>{''.join(linhas_compra_list)}</tbody>"
            "</table>"
            "</div>"
        )
        st.markdown(tabela_compras_html, unsafe_allow_html=True)

        pdf_bytes = gerar_pdf_lista_compras(df_compras, custo_total_compra)

        colunas_export = [
            "ID",
            "Item",
            "Categoria",
            "Unidade",
            "Status",
            "Qtd_Atual",
            "Est_Minimo",
            "Est_Maximo",
            "Qtd_Necessaria",
            "Melhor_Fornecedor",
            "Melhor_Preco_Unitario",
            "Custo_Estimado_Reposicao",
        ]
        df_export = df_compras[colunas_export].copy()
        csv_bytes = df_export.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="📄 Exportar Lista de Compras (PDF)",
                data=pdf_bytes,
                file_name=f"lista_compras_mazon_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        with col_dl2:
            st.download_button(
                label="📥 Exportar Lista de Compras (Planilha CSV / Excel)",
                data=csv_bytes,
                file_name=f"lista_compras_mazon_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
            )
