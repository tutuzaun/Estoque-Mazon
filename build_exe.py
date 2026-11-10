"""
Script de Empacotamento e Compilação (.EXE) - Restaurante Mazon
Utiliza o PyInstaller para empacotar o Streamlit + app.py + sheets_manager.py + style.css
e qualquer arquivo de logo existente (.jpeg, .jpg ou .png) em um executável Windows dentro da pasta dist/.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List

BASE_DIR = Path(__file__).resolve().parent
ENTRY_SCRIPT = BASE_DIR / "_launcher_mazon.py"
STREAMLIT_CONFIG_DIR = BASE_DIR / ".streamlit"
STREAMLIT_CONFIG_FILE = STREAMLIT_CONFIG_DIR / "config.toml"
EXE_NAME = "EstoqueMazon"

CANDIDATOS_LOGO = [
    "logo.jpeg",
    "logo.jpg",
    "logo.png",
    "WhatsApp Image 2026-11-09 at 17.04.39.jpeg",
]


def localizar_arquivos_logo() -> List[Path]:
    """Localiza todos os arquivos de logo (.jpeg, .jpg, .png) existentes na pasta do projeto."""
    encontrados: List[Path] = []
    for nome in CANDIDATOS_LOGO:
        caminho = BASE_DIR / nome
        if caminho.exists() and caminho not in encontrados:
            encontrados.append(caminho)

    for padrao in ["logo*.jpeg", "logo*.jpg", "logo*.png"]:
        for arq in BASE_DIR.glob(padrao):
            if arq.is_file() and arq not in encontrados:
                encontrados.append(arq)

    return encontrados


def garantir_config_streamlit() -> None:
    """
    Cria automaticamente a pasta .streamlit e o arquivo config.toml
    para evitar o prompt de e-mail na inicialização e desativar telemetria.
    """
    STREAMLIT_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    conteudo_config = (
        "[browser]\ngatherUsageStats = false\n\n"
        "[global]\ndevelopmentMode = false\n\n"
        "[theme]\n"
        'primaryColor = "#0B1B2B"\n'
        'backgroundColor = "#F4EDE3"\n'
        'secondaryBackgroundColor = "#FFFFFF"\n'
        'textColor = "#0B1B2B"\n'
    )
    STREAMLIT_CONFIG_FILE.write_text(conteudo_config, encoding="utf-8")
    print(f"[OK] Arquivo de configuração garantido em: {STREAMLIT_CONFIG_FILE}")


def criar_entrypoint_streamlit() -> None:
    """
    Cria o script wrapper necessário para que o binário compilado pelo PyInstaller
    invoque internamente o CLI do Streamlit apontando para o app.py empacotado.
    """
    codigo_launcher = '''# -*- coding: utf-8 -*-
import os
import sys
from streamlit.web import cli as stcli


def resolve_path(relative_path: str) -> str:
    """Resolve o caminho dos arquivos internos tanto no modo dev quanto no bundle PyInstaller (_MEIPASS)."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


if __name__ == "__main__":
    # Garante que o diretório de trabalho aponte para a pasta onde o .exe está localizado
    # para que os arquivos .streamlit/config.toml, logos e credentials.json sejam lidos normalmente
    if getattr(sys, "frozen", False):
        os.chdir(os.path.dirname(sys.executable))

    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

    app_path = resolve_path("app.py")
    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
    ]
    sys.exit(stcli.main())
'''
    ENTRY_SCRIPT.write_text(codigo_launcher, encoding="utf-8")


def verificar_instalar_pyinstaller() -> None:
    """Verifica se o PyInstaller está instalado no ambiente atual; caso contrário, instala."""
    try:
        import PyInstaller  # noqa: F401
        print("[OK] PyInstaller detectado no ambiente.")
    except ImportError:
        print("[INFO] PyInstaller não encontrado. Instalando automaticamente via pip...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])


def compilar_executavel() -> None:
    """Executa o PyInstaller com todos os submódulos, metadados, arquivos e logos do projeto Mazon."""
    verificar_instalar_pyinstaller()
    garantir_config_streamlit()
    criar_entrypoint_streamlit()

    import PyInstaller.__main__

    # Separador de caminho do --add-data no Windows é ';'
    sep = ";" if os.name == "nt" else ":"

    # Lista datas incluindo qualquer logo (.jpeg, .jpg ou .png) existente na pasta
    datas = [
        (BASE_DIR / "app.py", "."),
        (BASE_DIR / "sheets_manager.py", "."),
        (BASE_DIR / "style.css", "."),
        (STREAMLIT_CONFIG_FILE, ".streamlit"),
    ]

    logos_encontradas = localizar_arquivos_logo()
    for arq_logo in logos_encontradas:
        datas.append((arq_logo, "."))
        print(f"[OK] Logo detectada e adicionada em datas: {arq_logo.name}")

    argumentos_pyinstaller = [
        str(ENTRY_SCRIPT),
        f"--name={EXE_NAME}",
        "--onedir",
        "--noconfirm",
        "--clean",
    ]

    for origem, destino in datas:
        argumentos_pyinstaller.append(f"--add-data={origem}{sep}{destino}")

    argumentos_pyinstaller.extend(
        [
            # Coleta absolutamente todos os submódulos, binários, dados e metadados do Streamlit e Plotly
            "--collect-all=streamlit",
            "--copy-metadata=streamlit",
            "--collect-data=plotly",
            # Imports ocultos explícitos exigidos em tempo de execução pelo Streamlit
            "--hidden-import=streamlit",
            "--hidden-import=streamlit.runtime.scriptrunner.magic_funcs",
            "--hidden-import=streamlit.runtime.scriptrunner",
            "--hidden-import=streamlit.web.cli",
            "--hidden-import=pandas",
            "--hidden-import=plotly",
            "--hidden-import=gspread",
            "--hidden-import=google.oauth2.service_account",
            "--hidden-import=google_auth_oauthlib",
            "--hidden-import=googleapiclient.discovery",
            "--hidden-import=googleapiclient.http",
            "--hidden-import=openpyxl",
            "--hidden-import=PIL",
            "--hidden-import=reportlab",
            "--collect-all=reportlab",
        ]
    )

    print("======================================================================")
    print("  INICIANDO COMPILAÇÃO DO EXECUTÁVEL (.EXE) - RESTAURANTE MAZON")
    print("======================================================================")

    try:
        PyInstaller.__main__.run(argumentos_pyinstaller)

        # Copia também a pasta .streamlit e os arquivos de logo para dist/EstoqueMazon/
        dist_app_dir = BASE_DIR / "dist" / EXE_NAME
        if dist_app_dir.exists():
            dist_streamlit_dir = dist_app_dir / ".streamlit"
            dist_streamlit_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(STREAMLIT_CONFIG_FILE, dist_streamlit_dir / "config.toml")
            for arq_logo in logos_encontradas:
                shutil.copy2(arq_logo, dist_app_dir / arq_logo.name)

        print("\n======================================================================")
        print("  COMPILAÇÃO CONCLUÍDA COM SUCESSO!")
        print(f"  Executável gerado em: {dist_app_dir / (EXE_NAME + '.exe')}")
        print("======================================================================")
    finally:
        # Remove o script temporário de entrada após o build
        if ENTRY_SCRIPT.exists():
            ENTRY_SCRIPT.unlink()


if __name__ == "__main__":
    compilar_executavel()
