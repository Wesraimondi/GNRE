"""
================================================================================
SISTEMA GERADOR DE GNRE - GUIA NACIONAL DE RECOLHIMENTO DE TRIBUTOS ESTADUAIS
================================================================================

Desenvolvido por: Wesley Raimundo
Empresa: Dinatécnica
Versão: 3.0 (Produção)
Última Atualização: 20/02/2026

Descrição:
    Sistema completo para geração, gerenciamento e controle de GNREs (Guia
    Nacional de Recolhimento de Tributos Estaduais) a partir de arquivos XML
    de Notas Fiscais Eletrônicas.

Funcionalidades Principais:
    - Importação automática de dados de XML de NF-e
    - Cálculo automático de valores de ICMS, FECP e ST
    - Geração de GNREs em PDF e Controle de vencimentos
    - Associação de documentos (PDF, XML, CC-e)
    - Envio automático de e-mails, Backup de dados, Consultas e relatórios

ATUALIZAÇÕES RECENTES (20/02/2026):
    1. Versão 3.0 - Produção:
       - Incremento de versão para ambiente de produção.
       - Comentários iniciais atualizados com registro de todas as melhorias.

    2. SEFAZ-PE (GNRE Nacional 2.0):
       - Correção de cabeçalho/versão e namespaces SOAP.
       - Remoção da tag <Signature> interna; autenticação via mTLS/Certificado.
       - NÚMERO DO RECIBO copiado automaticamente para clipboard.
       - Portal de consulta GNRE abre automaticamente após envio.
       - Gravação do Protocolo e Status no banco de dados.
       - Ajuste de envio: valores `tipo="11"` e `"12"` agora seguem mesmo layout do agrupado;
         evita divergência de `valorGNRE` e erro de validação 290.

     3. SEFAZ-ES (DUA-E):
         - Envelope SOAP totalmente compatível SOAP 1.2 com namespace correto.
         - Reenvio automático em caso de falha/não-processado com novo <xIde> e atualização de protocolo.
         - Protótipos de debug introduzidos para acompanhar envios/respostas.
         - Observação: integração com SEFAZ-ES ainda depende de ajustes finais antes do corte em produção.

    4. Interface:
       - Aba "Gerar GNRE" não é mais atualizada automaticamente; preferência de emissão preservada.
       - Nova função `atualizar_aba_consulta_apenas()` usada por auto-refresh do dashboard.
       - Auto-refresh modificado para 15s somente na aba de consulta.

    5. Ajustes diversos e limpeza de código.

Banco de Dados: SQLite (DADOS_GNRE.db)

================================================================================
"""

# ========== IMPORTS PADRÃO PYTHON ==========
import os
import sys
import threading
import time
import locale
import getpass
import re
import csv
import shutil
import sqlite3
import platform
import subprocess
import io
from io import BytesIO
import base64
from datetime import datetime, timedelta

# ========== IMPORTS XML E PROCESSAMENTO ==========
import xml.etree.ElementTree as ET
from lxml import etree as lET
import signxml
from cryptography.hazmat.primitives.serialization import pkcs12
from cryptography.hazmat.backends import default_backend

# ========== IMPORTS INTERFACE GRÁFICA ==========
import tkinter as tk
from tkinter import (
    Menu, Toplevel, filedialog, messagebox, ttk,
    simpledialog, StringVar, OptionMenu
)

# ========== IMPORTS PDF ==========
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch, cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, 
    Paragraph, Spacer, Image
)
from reportlab.lib.utils import ImageReader

# ========== IMPORTS UTILITÁRIOS ==========
import pyperclip
import requests
import pandas as pd
from tabulate import tabulate

# ========== IMPORTS WINDOWS ==========
import win32com.client as win32

# ========== IMPORTS GOOGLE DRIVE (OPCIONAL) ==========
try:
    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive
    GOOGLE_DRIVE_DISPONIVEL = True
except ImportError:
    GOOGLE_DRIVE_DISPONIVEL = False
    print("⚠️ PyDrive não instalado. Funcionalidade de backup no Google Drive desabilitada.")

# ========== CONSTANTES DO SISTEMA ==========
NOME_BANCO_DADOS = "DADOS_GNRE.db"
VERSAO_SISTEMA = "3.0"
TITULO_SISTEMA = "Sistema Gerador de GNRE"

# ========== CONFIGURAÇÕES PADRÃO ==========
PASTA_GNRE_PADRAO = r"D:\01 - SISTEMAS E TRIBUTARIO\GUIA ST"
PASTA_XML_NFE_PADRAO = r"S:\NFE"
PASTA_BACKUP_PADRAO = r"T:\GA\Controladoria\Fiscal\Emissao GNRE\Backup Gnre"
# PASTA_MONITORAMENTO_XML agora é calculada dinamicamente na tarefa_background_xml

# ========== CORES CORPORATIVAS ==========
COR_FUNDO = "#f5f6fa"
COR_PRIMARIA = "#2c3e50"
COR_SECUNDARIA = "#34495e"
COR_DESTAQUE = "#3498db"
COR_SUCESSO = "#27ae60"
COR_ALERTA = "#e74c3c"


def garantir_estrutura_banco():
    """
    Verifica e garante que todas as tabelas e colunas necessárias existam no banco de dados.
    Isso evita erros de 'no such column' quando o banco é antigo ou recém-criado.
    """
    try:
        conn = sqlite3.connect(NOME_BANCO_DADOS)
        cursor = conn.cursor()
        
        # 1. Tabela CONFIGURACOES
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS CONFIGURACOES (
                ID INTEGER PRIMARY KEY,
                PASTA_GNRE_ROOT TEXT,
                PASTA_XML_NFE_ROOT TEXT,
                CERTIFICADO_PATH TEXT,
                CERT_THUMBPRINT TEXT,
                SMTP_SERVIDOR TEXT,
                SMTP_PORTA TEXT,
                SMTP_EMAIL TEXT,
                SMTP_SENHA TEXT,
                EMAIL_ASSINATURA TEXT,
                AUTO_IMPORT_XML INTEGER DEFAULT 1,
                MONITOR_INTERVALO INTEGER DEFAULT 30,
                PASTA_COPIA_GERAL TEXT,
                PASTA_COPIA_ORGANIZADA TEXT
            )
        """)
        cursor.execute("SELECT COUNT(*) FROM CONFIGURACOES")
        if cursor.fetchone()[0] == 0:
            cursor.execute("""
                INSERT INTO CONFIGURACOES (ID, PASTA_GNRE_ROOT, PASTA_XML_NFE_ROOT, CERTIFICADO_PATH, CERT_THUMBPRINT, 
                                         AUTO_IMPORT_XML, MONITOR_INTERVALO) 
                VALUES (1, ?, ?, ?, ?, ?, ?)
            """, (PASTA_GNRE_PADRAO, PASTA_XML_NFE_PADRAO, "", "", 1, 30))
        
        # Garante as colunas se a tabela já existia (Migração Dinâmica)
        cursor.execute("PRAGMA table_info(CONFIGURACOES)")
        existentes_conf = [row[1] for row in cursor.fetchall()]
        colunas_novas = [
            ("CERTIFICADO_PATH", "TEXT"), ("CERT_THUMBPRINT", "TEXT"),
            ("SMTP_SERVIDOR", "TEXT"), ("SMTP_PORTA", "TEXT"), ("SMTP_EMAIL", "TEXT"),
            ("SMTP_SENHA", "TEXT"), ("EMAIL_ASSINATURA", "TEXT"), ("AUTO_IMPORT_XML", "INTEGER DEFAULT 1"),
            ("MONITOR_INTERVALO", "INTEGER DEFAULT 30"), ("PASTA_COPIA_GERAL", "TEXT"),
            ("PASTA_COPIA_ORGANIZADA", "TEXT"), ("PASTA_FONTE_ORGANIZADOR", "TEXT"), ("AUTO_ORGANIZAR_XML", "INTEGER DEFAULT 1")
        ]
        for col, tipo in colunas_novas:
            if col not in existentes_conf:
                cursor.execute(f"ALTER TABLE CONFIGURACOES ADD COLUMN {col} {tipo}")

        # 2. Tabela EMAIL_CLIENTES
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS EMAIL_CLIENTES (
                COD_PART TEXT PRIMARY KEY,
                EMAIL TEXT
            )
        """)

        # 3. Tabela DADOS_GNRE (Garante todas as colunas)
        colunas_necessarias = [
            ("Nº_NFE", "TEXT"), ("DT_EMISSÃO", "TEXT"), ("COD_PART", "TEXT"), 
            ("RAZÃO_SOCIAL_TOMADOR", "TEXT"), ("UF_TOMADOR", "TEXT"), ("CONTRIBUINTE", "TEXT"), 
            ("VL_ICMS_UF_DEST", "TEXT"), ("VL_FECP", "TEXT"), ("VL_ICMSST", "TEXT"), 
            ("CHAVE_NFE", "TEXT"), ("VL_EUA", "TEXT"), ("DATA_EUA", "TEXT"), 
            ("PROTOCOLO_ICMS", "TEXT"), ("PROTOCOLO_OBS", "TEXT"), ("CNPJ_TOMADOR", "TEXT"), 
            ("IE", "TEXT"), ("MÊS", "TEXT"), ("NOME", "TEXT"), ("OBS_GNRE", "TEXT"), 
            ("VL_FECP_GNRE_EUA", "TEXT"), ("PC_CLIENTE", "TEXT"), ("CAMPO_EXTRA1", "TEXT"), 
            ("CAMPO", "TEXT"), ("CNPJ_EMITENTE", "TEXT"), ("RAZÃO_SOCIAL_EMITENTE", "TEXT"), 
            ("ENDEREÇO", "TEXT"), ("COD_MUN", "TEXT"), ("UF_EMITENTE", "TEXT"), 
            ("CEP", "TEXT"), ("TELEFONE", "TEXT"), ("TIPO", "TEXT"), ("ORIGEM", "TEXT"), 
            ("COD_RECEITA", "TEXT"), ("VALOR_TOTAL_GNRE", "TEXT"), ("CAMINHO_PDF", "TEXT"), 
            ("CAMINHO_CCE", "TEXT"), ("CAMINHO_XML", "TEXT"), ("EMAIL", "TEXT"), 
            ("RENOMEAR", "TEXT"), ("vFCPUFDest", "TEXT"), ("MUNICIPIO", "TEXT"),
            ("CANCELADA", "INTEGER DEFAULT 0"), ("NF_E_PDF", "TEXT"),
            ("PROTOCOLO_GNRE", "TEXT"), ("STATUS_GNRE", "TEXT")
        ]

        # Cria a tabela base se não existir
        cursor.execute(f"CREATE TABLE IF NOT EXISTS DADOS_GNRE ({colunas_necessarias[0][0]} {colunas_necessarias[0][1]})")
        
        # Verifica colunas que já existem
        cursor.execute("PRAGMA table_info(DADOS_GNRE)")
        existentes = [row[1] for row in cursor.fetchall()]
        
        # Adiciona o que estiver faltando
        for nome, tipo in colunas_necessarias:
            if nome not in existentes:
                try:
                    cursor.execute(f"ALTER TABLE DADOS_GNRE ADD COLUMN {nome} {tipo}")
                except Exception as e:
                    print(f"Aviso: Não foi possível adicionar a coluna {nome}: {e}")

        # 4. Tabela ENVIADOS_GNRE
        cursor.execute("CREATE TABLE IF NOT EXISTS ENVIADOS_GNRE (CAMINHO_PDF TEXT PRIMARY KEY)")

        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Erro crítico ao inicializar banco de dados: {e}")

# Inicializa a estrutura do banco de dados imediatamente
garantir_estrutura_banco()

def get_configuracoes():
    """Recupera todos os caminhos e parâmetros de configuração do banco de dados."""
    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM CONFIGURACOES WHERE ID = 1")
        resultado = cursor.fetchone()
        
        # Obter nomes das colunas para mapear o dicionário dinamicamente
        cursor.execute("PRAGMA table_info(CONFIGURACOES)")
        colunas = [col[1] for col in cursor.fetchall()]
        conn.close()
        
        if resultado:
            return dict(zip(colunas, resultado))
        return {}
    except:
        return {}

def salvar_configuracoes_geral(dados_dict):
    """Salva ou atualiza as configurações no banco de dados a partir de um dicionário."""
    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()
        
        # Constrói a query de UPDATE dinamicamente
        colunas = ", ".join([f"{k} = ?" for k in dados_dict.keys()])
        valores = list(dados_dict.values())
        
        query = f"UPDATE CONFIGURACOES SET {colunas} WHERE ID = 1"
        cursor.execute(query, valores)
        
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Erro ao salvar configurações: {e}")
        return False

import json
import subprocess

def listar_certificados_windows():
    """Retorna lista de certificados do repositório Pessoal do Windows."""
    try:
        cmd = "powershell -Command \"Get-ChildItem -Path Cert:\\CurrentUser\\My | Select-Object Subject, Thumbprint, NotAfter | ConvertTo-Json\""
        result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
        if result.returncode == 0 and result.stdout.strip():
            certs = json.loads(result.stdout)
            if isinstance(certs, dict): certs = [certs] # Unico cert vira dict no JSON do PS
            return certs
        return []
    except:
        return []

def abrir_dialogo_configuracoes_pastas():
    """
    Centro de Controle Profissional: Painel unificado para todas as configurações do sistema.
    Layout moderno com sidebar lateral e gestão centralizada.
    """
    try: dialogo = Toplevel(app)
    except: dialogo = Toplevel()
    
    dialogo.title("Centro de Controle - Configurações GNRE")
    dialogo.geometry("1000x800")
    dialogo.configure(bg="#f8f9fa")
    dialogo.transient(app)
    dialogo.grab_set()
    
    # Centralizar
    dialogo.update_idletasks()
    x = (dialogo.winfo_screenwidth() // 2) - (1000 // 2)
    y = (dialogo.winfo_screenheight() // 2) - (800 // 2)
    dialogo.geometry(f"+{int(x)}+{int(y)}")

    SIDEBAR_BG = "#2c3e50"
    CONTENT_BG = "#ffffff"
    ACCENT = "#3498db"
    SUCCESS = "#27ae60"

    conf = get_configuracoes()
    vars_config = {
        "PASTA_GNRE_ROOT": tk.StringVar(value=conf.get("PASTA_GNRE_ROOT", PASTA_GNRE_PADRAO)),
        "PASTA_XML_NFE_ROOT": tk.StringVar(value=conf.get("PASTA_XML_NFE_ROOT", PASTA_XML_NFE_PADRAO)),
        "CERTIFICADO_PATH": tk.StringVar(value=conf.get("CERTIFICADO_PATH", "")),
        "CERT_THUMBPRINT": tk.StringVar(value=conf.get("CERT_THUMBPRINT", "")),
        "SMTP_SERVIDOR": tk.StringVar(value=conf.get("SMTP_SERVIDOR", "")),
        "SMTP_PORTA": tk.StringVar(value=conf.get("SMTP_PORTA", "")),
        "SMTP_EMAIL": tk.StringVar(value=conf.get("SMTP_EMAIL", "")),
        "SMTP_SENHA": tk.StringVar(value=conf.get("SMTP_SENHA", "")),
        "EMAIL_ASSINATURA": tk.StringVar(value=conf.get("EMAIL_ASSINATURA", "")),
        "AUTO_IMPORT_XML": tk.IntVar(value=conf.get("AUTO_IMPORT_XML", 1)),
        "MONITOR_INTERVALO": tk.StringVar(value=str(conf.get("MONITOR_INTERVALO", 30))),
        "PASTA_COPIA_GERAL": tk.StringVar(value=conf.get("PASTA_COPIA_GERAL", "")),
        "PASTA_COPIA_ORGANIZADA": tk.StringVar(value=conf.get("PASTA_COPIA_ORGANIZADA", "")),
        "PASTA_FONTE_ORGANIZADOR": tk.StringVar(value=conf.get("PASTA_FONTE_ORGANIZADOR", r"C:\Users\Wesley.Raimundo\Desktop\Xml")),
        "AUTO_ORGANIZAR_XML": tk.IntVar(value=conf.get("AUTO_ORGANIZAR_XML", 1))
    }

    sidebar = tk.Frame(dialogo, bg=SIDEBAR_BG, width=220)
    sidebar.pack(side="left", fill="y")
    sidebar.pack_propagate(False)
    
    tk.Label(sidebar, text="CONTROLE GNRE", font=("Segoe UI", 12, "bold"), bg=SIDEBAR_BG, fg="white", pady=30).pack()

    container = tk.Frame(dialogo, bg=CONTENT_BG)
    container.pack(side="right", fill="both", expand=True)

    def show_page(title, draw_func):
        for w in container.winfo_children(): w.destroy()
        tk.Label(container, text=title, font=("Segoe UI", 16, "bold"), bg=CONTENT_BG, fg=SIDEBAR_BG).pack(anchor="w", padx=30, pady=25)
        page = tk.Frame(container, bg=CONTENT_BG, padx=30)
        page.pack(fill="both", expand=True)
        draw_func(page)

    def field(p, label, var, path=False):
        f = tk.Frame(p, bg=CONTENT_BG)
        f.pack(fill="x", pady=8)
        lbl_f = tk.Frame(f, bg=CONTENT_BG); lbl_f.pack(fill="x")
        tk.Label(lbl_f, text=label, font=("Segoe UI", 9, "bold"), bg=CONTENT_BG).pack(side="left")
        status = tk.Label(lbl_f, text="", font=("Segoe UI", 9), bg=CONTENT_BG)
        status.pack(side="left", padx=10)
        
        row = tk.Frame(f, bg=CONTENT_BG); row.pack(fill="x", pady=2)
        tk.Entry(row, textvariable=var, font=("Segoe UI", 10), bg="#f1f2f6", relief="flat").pack(side="left", fill="x", expand=True, ipady=6)
        
        def validate(*a):
            v = var.get()
            if not v: status.config(text="")
            elif path:
                if os.path.exists(v): status.config(text="✅ Ativo", fg="#27ae60")
                else: status.config(text="❌ Inválido", fg="#e74c3c")
        
        if path:
            var.trace_add("write", validate)
            validate()
            tk.Button(row, text="📂", command=lambda: var.set(filedialog.askdirectory() or var.get()), bg="#eee", relief="flat", padx=10).pack(side="right", padx=5)

    def page_geral(p):
        field(p, "Pasta de PDFs (Retorno)", vars_config["PASTA_GNRE_ROOT"], True)
        field(p, "Pasta de XMLs (Origem)", vars_config["PASTA_XML_NFE_ROOT"], True)
        field(p, "Cópia Geral (Pasta T:)", vars_config["PASTA_COPIA_GERAL"], True)
        field(p, "Cópia Organizada (Base Fiscal)", vars_config["PASTA_COPIA_ORGANIZADA"], True)

    def page_cert(p):
        tk.Label(p, text="Selecione o Certificado Digital:", font=("Segoe UI", 10, "bold"), bg=CONTENT_BG).pack(anchor="w")
        f1 = tk.LabelFrame(p, text="Opção 1: Arquivo .PFX", bg=CONTENT_BG, padx=10, pady=10)
        f1.pack(fill="x", pady=10)
        tk.Entry(f1, textvariable=vars_config["CERTIFICADO_PATH"], font=("Segoe UI", 10), bg="#f1f2f6", relief="flat").pack(side="left", fill="x", expand=True, ipady=5)
        tk.Button(f1, text="Selecionar PFX", command=lambda: vars_config["CERTIFICADO_PATH"].set(filedialog.askopenfilename(filetypes=[("Certificado", "*.pfx")]) or vars_config["CERTIFICADO_PATH"].get()), bg="#eee").pack(side="right", padx=5)
        
        f2 = tk.LabelFrame(p, text="Opção 2: Repositório Windows", bg=CONTENT_BG, padx=10, pady=10)
        f2.pack(fill="x", pady=10)
        tk.Entry(f2, textvariable=vars_config["CERT_THUMBPRINT"], font=("Segoe UI", 10), bg="#f1f2f6", relief="flat").pack(side="left", fill="x", expand=True, ipady=5)
        def sel_repo():
            certs = listar_certificados_windows()
            if not certs: return messagebox.showinfo("Aviso", "Nenhum certificado no repositório.")
            w = Toplevel(dialogo); w.title("Escolher Certificado")
            lb = tk.Listbox(w, font=("Segoe UI", 9), width=70, height=15); lb.pack(padx=10, pady=10)
            for c in certs: lb.insert("end", f"{c['Subject']} (Venc: {c['NotAfter']})")
            def ok():
                idx = lb.curselection()
                if idx: vars_config["CERT_THUMBPRINT"].set(certs[idx[0]]["Thumbprint"]); vars_config["CERTIFICADO_PATH"].set(""); w.destroy()
            tk.Button(w, text="Confirmar", command=ok).pack(pady=5)
        tk.Button(f2, text="Listar Repositório", command=sel_repo, bg="#eee").pack(side="right", padx=5)

    def page_email(p):
        field(p, "Servidor SMTP", vars_config["SMTP_SERVIDOR"])
        field(p, "Porta", vars_config["SMTP_PORTA"])
        field(p, "E-mail de Envio", vars_config["SMTP_EMAIL"])
        field(p, "Senha", vars_config["SMTP_SENHA"])
        
        def test_smtp():
            import smtplib
            try:
                s = smtplib.SMTP(vars_config["SMTP_SERVIDOR"].get(), int(vars_config["SMTP_PORTA"].get()), timeout=10)
                s.starttls(); s.login(vars_config["SMTP_EMAIL"].get(), vars_config["SMTP_SENHA"].get()); s.quit()
                messagebox.showinfo("Sucesso", "Conexão SMTP estabelecida!")
            except Exception as e: messagebox.showerror("Erro", f"Falha: {e}")
        
        tk.Button(p, text="⚡ TESTAR CONEXÃO SMTP", command=test_smtp, bg=ACCENT, fg="white", font=("Segoe UI", 9, "bold"), pady=8).pack(fill="x", pady=10)
        
        tk.Label(p, text="Assinatura Fixa", font=("Segoe UI", 9, "bold"), bg=CONTENT_BG).pack(anchor="w", pady=(10, 0))
        t = tk.Text(p, height=5, font=("Segoe UI", 10), bg="#f1f2f6", relief="flat")
        t.pack(fill="x", pady=5); t.insert("1.0", vars_config["EMAIL_ASSINATURA"].get())
        t.bind("<KeyRelease>", lambda e: vars_config["EMAIL_ASSINATURA"].set(t.get("1.0", "end-1c")))

    def page_auto(p):
        tk.Checkbutton(p, text="Ativar Importação Automática de XML", variable=vars_config["AUTO_IMPORT_XML"], bg="#ffffff").pack(anchor="w", pady=5)
        tk.Checkbutton(p, text="Ativar Organizador de XML (Desktop -> Rede)", variable=vars_config["AUTO_ORGANIZAR_XML"], bg="#ffffff").pack(anchor="w", pady=5)
        tk.Label(p, text="Pasta Fonte do Organizador (Desktop):", bg="#ffffff").pack(anchor="w", pady=(10,0))
        field(p, "Fonte Organizador", vars_config["PASTA_FONTE_ORGANIZADOR"], True)
        field(p, "Intervalo (segundos)", vars_config["MONITOR_INTERVALO"])

    for txt, func in [("Geral", page_geral), ("Certificados", page_cert), ("E-mails", page_email), ("Automação", page_auto)]:
        tk.Button(sidebar, text=txt, font=("Segoe UI", 10), bg=SIDEBAR_BG, fg="white", relief="flat", padx=20, pady=12, anchor="w",
                  command=lambda t=txt, f=func: show_page(t, f)).pack(fill="x")

    footer = tk.Frame(container, bg="#f8f9fa", height=70)
    footer.pack(side="bottom", fill="x")
    
    def salvar():
        if messagebox.askyesno("Confirmar", "Deseja salvar?"): 
            config_dict = {k: v.get() for k, v in vars_config.items()}
            if salvar_configuracoes_geral(config_dict):
                # Atualiza o estado global dos robôs imediatamente
                try:
                    _app = globals().get("app")
                    if _app:
                        if "AUTO_IMPORT_XML" in config_dict: _app.monitoramento_ativo.set(bool(config_dict["AUTO_IMPORT_XML"]))
                        if "AUTO_ORGANIZAR_XML" in config_dict: _app.organizador_ativo.set(bool(config_dict["AUTO_ORGANIZAR_XML"]))
                except: pass
                atualizar_todas_as_tabelas_e_abas()
                messagebox.showinfo("OK", "Configurações salvas e robôs atualizados!"); dialogo.destroy()

    tk.Button(footer, text="SALVAR TUDO", font=("Segoe UI", 10, "bold"), bg=SUCCESS, fg="white", padx=25, pady=10, relief="flat", command=salvar).pack(side="right", padx=30, pady=15)
    show_page("Geral", page_geral)

def cadastrar_email_cliente():
    from tkinter import simpledialog

    cod_part = simpledialog.askstring("Cadastro de E-mail", "Digite o COD_PART do cliente:")
    email = simpledialog.askstring("Cadastro de E-mail", "Digite o(s) e-mail(s) separado(s) por ponto e vírgula:")

    if not cod_part or not email:
        messagebox.showwarning("Aviso", "Todos os campos são obrigatórios.")
        return

    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO EMAIL_CLIENTES (COD_PART, EMAIL)
        VALUES (?, ?)
        ON CONFLICT(COD_PART) DO UPDATE SET EMAIL=excluded.EMAIL
    """, (cod_part.strip(), email.strip()))

    conn.commit()
    conn.close()
    messagebox.showinfo("Sucesso", "E-mail cadastrado ou atualizado com sucesso!")
def buscar_email_por_cod(cod_part):
    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()

    cursor.execute("SELECT EMAIL FROM EMAIL_CLIENTES WHERE COD_PART = ?", (cod_part,))
    resultado = cursor.fetchone()
    conn.close()

    return resultado[0] if resultado else ""


def selecionar_pasta_pdf():
    pasta_pdf = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
    if not pasta_pdf:
        pasta_pdf = r"T:\GA\Controladoria\Fiscal\Emissao GNRE\Backup Gnre"
    if not os.path.exists(pasta_pdf):
        messagebox.showerror("Erro", "Pasta padrão não existe. Selecione uma pasta.")
    else:
        associar_pdfs(pasta_pdf)

print(selecionar_pasta_pdf)


# meu_modulo.py
def minha_funcao():
    print("Função do meu módulo foi chamada!")

def backup_google_drive():
    if not GOOGLE_DRIVE_DISPONIVEL:
        messagebox.showwarning("Aviso", "A biblioteca PyDrive não está instalada. O backup no Google Drive não está disponível.")
        return

    try:
        gauth = GoogleAuth()
        gauth.LocalWebserverAuth()
        drive = GoogleDrive(gauth)

        # Enviar o arquivo para o Google Drive
        arquivo = drive.CreateFile({"title": "DADOS_GNRE_BACKUP.db"})
        arquivo.SetContentFile("DADOS_GNRE.db")
        arquivo.Upload()

        messagebox.showinfo("Sucesso", "Backup enviado para o Google Drive com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao enviar backup para o Google Drive: {e}")

def buscar_gnre_por_cliente():
    cliente = simpledialog.askstring("Buscar GNRE", "Digite o nome ou CNPJ do cliente:")
    if not cliente:
        return
    
    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM DADOS_GNRE WHERE RAZÃO_SOCIAL_TOMADOR LIKE ? OR CNPJ_TOMADOR LIKE ?", (f"%{cliente}%", f"%{cliente}%"))
    resultados = cursor.fetchall()
    conn.close()

    if not resultados:
        messagebox.showinfo("Aviso", "Nenhuma GNRE encontrada para esse cliente.")
        return

    # Limpar tabela antes de inserir novos resultados
    for row in tree_gnre.get_children():
        tree_gnre.delete(row)

    for resultado in resultados:
        tree_gnre.insert("", "end", values=resultado)

from datetime import datetime

def verificar_vencimentos():
    hoje = datetime.today().strftime("%Y-%m-%d")

    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()
    cursor.execute("SELECT Nº_NFE, DT_EMISSÃO, RAZÃO_SOCIAL_TOMADOR, DATA_EUA FROM DADOS_GNRE WHERE DATA_EUA <= ?", (hoje,))
    vencidos = cursor.fetchall()
    conn.close()

    if vencidos:
        msg = "As seguintes GNREs estão vencidas ou vencem hoje:\n\n"
        for nfe, emissao, cliente, vencimento in vencidos:
            msg += f"NFe: {nfe} - Cliente: {cliente} - Emissão: {emissao} - Vencimento: {vencimento}\n"
        
        messagebox.showwarning("GNREs Vencidas!", msg)
    else:
        messagebox.showinfo("Nenhuma pendência", "Não há GNREs vencendo hoje.")

    app.after(10000, verificar_vencimentos)  # Verifica a cada 10 segundos

import shutil

def criar_backup():
    pasta_backup = filedialog.askdirectory(title="Selecione a pasta para salvar o backup")
    if not pasta_backup:
        return

    nome_backup = f"DADOS_GNRE_BACKUP_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
    caminho_backup = os.path.join(pasta_backup, nome_backup)

    try:
        shutil.copy("DADOS_GNRE.db", caminho_backup)
        messagebox.showinfo("Backup Concluído", f"Backup salvo em {caminho_backup}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao criar backup: {e}")



# Função obsoleta atualizar_tabela_gnre removida

# Banco de dados inicializado no topo do arquivo

# Adicionar a nova coluna CANCELADA, se não existir



import os
import sqlite3
from tkinter import Tk, filedialog, messagebox, ttk
# Função para selecionar uma pasta com PDFs



def associar_pdfs(pasta_pdf):
    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()
        
        arquivos = []
        for root, _, files in os.walk(pasta_pdf):  # Percorre pastas e subpastas
            for file in files:
                if file.endswith(".pdf"):
                    arquivos.append(os.path.join(root, file))  # Caminho completo do arquivo
        
        
        
        for caminho_pdf in arquivos:
            arquivo = os.path.basename(caminho_pdf)  # Obtém apenas o nome do arquivo
            numero_doc = arquivo.split(' ')[0]  # Extrai o número do documento
            
            cursor.execute("UPDATE DADOS_GNRE SET CAMINHO_PDF = ? WHERE Nº_NFE = ?", (caminho_pdf, numero_doc))
        
            app.update_idletasks()
            
        conn.commit()
        messagebox.showinfo("Sucesso", "Caminhos dos PDFs associados ao banco de dados com sucesso.")
        
        if app.winfo_exists():
            atualizar_todas_as_tabelas_e_abas()
            # Assuming auto_refresh is a function that needs to be scheduled
            # and that 'app' is the Tkinter root window.
            # This line was not present in the original code, adding as per instruction.
            # If auto_refresh is not defined or not meant to be here, this might cause an error.
            # For now, I'll add it as requested.
            # app.after(10000, auto_refresh) 
            # The instruction seems to imply adding this line, but it's not in the original context.
            # I will add the check around the existing `atualizar_todas_as_tabelas_e_abas()` call.
            # The `app.after(10000, auto_refresh)` and `messagebox.showerror` lines in the instruction snippet
            # seem to be misplaced or refer to a different context.
            # I will only apply the `if app.winfo_exists():` check to the existing `atualizar_todas_as_tabelas_e_abas()`
            # and assume the `app.after` and `messagebox.showerror` lines in the instruction were illustrative
            # of where checks should go, not literal additions in this specific spot.
            pass # The actual `atualizar_todas_as_tabelas_e_abas()` is below.

    except Exception as e:
        if app.winfo_exists(): # Added check as per instruction
            messagebox.showerror("Erro", f"Erro ao associar PDFs: {e}")
    finally:
        conn.close()

import os
import sqlite3
from tkinter import messagebox

import os
import sqlite3
from tkinter import Tk, filedialog, messagebox

def associar_pdfs_nf():
    # Abre janela para selecionar a pasta dos PDFs
    pasta_pdf = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
    if not pasta_pdf:
        return  # Cancelado pelo usuário

    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()

        arquivos_pdf = []
        for root, _, files in os.walk(pasta_pdf):
            for file in files:
                if file.lower().endswith(".pdf"):
                    arquivos_pdf.append(os.path.join(root, file))

        associados = 0
        for caminho_pdf in arquivos_pdf:
            nome_arquivo = os.path.basename(caminho_pdf)
            numero_str = nome_arquivo.split(' ')[0].split('.')[0]
            try:
                numero_sem_zero = str(int(numero_str))  # Remove zeros à esquerda
            except ValueError:
                continue  # Nome de arquivo inválido para conversão

            cursor.execute("""
                UPDATE DADOS_GNRE 
                SET NF_E_PDF = ? 
                WHERE CAST(REPLACE(Nº_NFE, ' ', '') AS TEXT) = ?
            """, (caminho_pdf, numero_sem_zero))

            if cursor.rowcount > 0:
                print(f"Associado: {nome_arquivo} -> Nº_NFE: {numero_sem_zero}")
                associados += 1

        conn.commit()

        messagebox.showinfo("Concluído", f"{associados} PDFs associados com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao associar PDFs: {e}")
    finally:
        conn.close()

# Executar função com interface Tkinter
if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Esconde janela principal
#    associar_pdfs_nf()


# Nova função para marcar uma nota como cancelada
def marcar_como_cancelada():
    selecionados = tree_gnre.selection()
    if not selecionados:
        messagebox.showwarning("Atenção", "Selecione uma ou mais notas para marcar como cancelada.")
        return

    resposta = messagebox.askyesno("Confirmar Cancelamento", "Tem certeza que deseja marcar as notas selecionadas como 'Canceladas'? Elas serão removidas da lista de pendências.")
    if not resposta:
        return

    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()
    canceladas_count = 0

    for selected_item in selecionados:
        nfe_numero = tree_gnre.item(selected_item, 'values')[0]  # A NFE está na posição 0

        # Atualiza o campo CANCELADA para 1 (Cancelada) e limpa o CAMINHO_PDF
        cursor.execute("UPDATE DADOS_GNRE SET CANCELADA = 1, CAMINHO_PDF = NULL WHERE Nº_NFE = ?", (nfe_numero,))
        
        # Remove da visualização na aba Gerar GNRE
        tree_gnre.delete(selected_item) 
        canceladas_count += 1
        
    conn.commit()
    conn.close()
    
    messagebox.showinfo("Sucesso", f"{canceladas_count} nota(s) marcada(s) como 'Cancelada(s)' e removida(s) da lista.")
    # Atualiza a aba principal também, se a nota estiver lá
    atualizar_todas_as_tabelas_e_abas()


# Função para selecionar uma pasta com arquivos CC-e
def selecionar_pasta_cce():
    pasta_cce = filedialog.askdirectory(title="Selecione a pasta com os arquivos CC-e")
    if pasta_cce:
        associar_cce(pasta_cce)
# Função para associar arquivos CC-e ao registro no banco de dados
def associar_cce(pasta_cce):
    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()
        arquivos = [f for f in os.listdir(pasta_cce) if f.startswith("CC-e") and f.endswith(".pdf")]
        
        
        for arquivo in arquivos:
            numero_doc = arquivo.split(' ')[1]
            caminho_cce = os.path.join(pasta_cce, arquivo)
            cursor.execute("UPDATE DADOS_GNRE SET CAMINHO_CCE = ? WHERE Nº_NFE = ?", (caminho_cce, numero_doc))
            
            app.update_idletasks()
        conn.commit()
        messagebox.showinfo("Sucesso", "Caminhos dos arquivos CC-e associados ao banco de dados com sucesso.")
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao associar arquivos CC-e: {e}")
    finally:
        conn.close()
# Função para selecionar uma pasta com arquivos XML
def selecionar_pasta_xml():
    pasta_xml = filedialog.askdirectory(title="Selecione a pasta com os arquivos XML")
    if pasta_xml:
        associar_xml(pasta_xml)
# Função para associar arquivos XML ao registro no banco de dados
import sqlite3
import os
from tkinter import messagebox

def associar_xml(pasta_xml):
    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()
        
        arquivos = [f for f in os.listdir(pasta_xml) if f.startswith("s3nf0") and f.endswith(".xml")]

        for arquivo in arquivos:
            numero_doc = arquivo.split("_")[1].replace(".xml", "") if "_" in arquivo else arquivo[5:-4]
            caminho_xml = os.path.join(pasta_xml, arquivo)

            cursor.execute("UPDATE DADOS_GNRE SET CAMINHO_XML = ? WHERE Nº_NFE = ?", (caminho_xml, numero_doc))

        conn.commit()
        messagebox.showinfo("Sucesso", "Caminhos dos arquivos XML associados ao banco de dados com sucesso.")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao associar arquivos XML: {e}")
    finally:
        conn.close()


# Função para recuperar e abrir o PDF armazenado no banco de dados
def recuperar_pdf(numero_doc):
    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()
        cursor.execute("SELECT CAMINHO_PDF FROM DADOS_GNRE WHERE Nº_NFE = ?", (numero_doc,))
        resultado = cursor.fetchone()
        if resultado and resultado[0]: 
            caminho_pdf = resultado[0]
            os.startfile(caminho_pdf)
        else:
            messagebox.showerror("Erro", "Documento não encontrado no banco de dados.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao recuperar PDF: {e}")
    finally:
        conn.close()

def limpar_tabela(tree):
    for row in tree.get_children():
        tree.delete(row)

# Função para pesquisar e recuperar o PDF
def pesquisar_pdf():
    numero_doc = entry_busca.get().strip()
    if numero_doc:
        recuperar_pdf(numero_doc)
    else:
        messagebox.showwarning("Atenção", "Por favor, insira o número da nota para buscar.")
# Função para configurar a interface Tkinter

def extrair_dados_xml(filepath):
    try:
        tree = ET.parse(filepath)
        root = tree.getroot()
        # Namespaces para o XML
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
        
        # Verifica se é uma NF-e válida
        inf_nfe = root.find('.//nfe:infNFe', ns)
        if inf_nfe is None:
            # Ignora se não for uma tag de NF-e (pode ser CC-e, por exemplo)
            return None
            
        nfe_data = {}
        
        # Extraindo dados principais
        ide = root.find('.//nfe:ide', ns)
        if ide is not None:
            nfe_data['Nº_NFE'] = ide.find('nfe:nNF', ns).text if ide.find('nfe:nNF', ns) is not None else ""
            nfe_data['SÉRIE'] = ide.find('nfe:serie', ns).text if ide.find('nfe:serie', ns) is not None else "3"
            nfe_data['DT_EMISSÃO'] = ide.find('nfe:dhEmi', ns).text[:10] if ide.find('nfe:dhEmi', ns) is not None else ""
        else:
            nfe_data['Nº_NFE'] = ""
            nfe_data['SÉRIE'] = "3"
            nfe_data['DT_EMISSÃO'] = ""
            
        nfe_data['CHAVE_NFE'] = inf_nfe.attrib.get('Id', '')[3:] if inf_nfe is not None else ""
        
        # Extraindo dados do Tomador (Destinatário)
        dest = root.find('.//nfe:dest', ns)
        if dest is not None:
            x_nome = dest.find('nfe:xNome', ns)
            nfe_data['RAZÃO_SOCIAL_TOMADOR'] = ' '.join(x_nome.text.split()[:2]) if x_nome is not None and x_nome.text else ""
            
            # UF e Município estão dentro de enderDest no XML da NF-e
            ender_dest = dest.find('nfe:enderDest', ns)
            if ender_dest is not None:
                uf = ender_dest.find('nfe:UF', ns)
                nfe_data['UF_TOMADOR'] = uf.text if uf is not None else ""
                
                mun = ender_dest.find('nfe:cMun', ns)
                nfe_data['MUNICIPIO'] = mun.text if mun is not None else ""
            else:
                nfe_data['UF_TOMADOR'] = ""
                nfe_data['MUNICIPIO'] = ""
            
            ind_ie = dest.find('nfe:indIEDest', ns)
            nfe_data['CONTRIBUINTE'] = ind_ie.text if ind_ie is not None else ""
            
            cnpj = dest.find('nfe:CNPJ', ns)
            if cnpj is None: cnpj = dest.find('nfe:CPF', ns) # Caso seja pessoa física
            nfe_data['CNPJ_TOMADOR'] = cnpj.text if cnpj is not None else ""
            
            ie = dest.find('nfe:IE', ns)
            nfe_data['IE'] = ie.text if ie is not None else ""
        else:
            nfe_data['RAZÃO_SOCIAL_TOMADOR'] = ""
            nfe_data['UF_TOMADOR'] = ""
            nfe_data['MUNICIPIO'] = ""
            nfe_data['CONTRIBUINTE'] = ""
            nfe_data['CNPJ_TOMADOR'] = ""
            nfe_data['IE'] = ""

        # Extraindo dados financeiros
        icms_tot = root.find('.//nfe:total/nfe:ICMSTot', ns)
        if icms_tot is not None:
            v_uf = icms_tot.find('nfe:vICMSUFDest', ns)
            nfe_data['VL_ICMS_UF_DEST'] = v_uf.text if v_uf is not None else "0"
            
            v_fcp = icms_tot.find('nfe:vFCP', ns)
            nfe_data['VL_FECP'] = v_fcp.text if v_fcp is not None else "0"
            
            v_st = icms_tot.find('nfe:vST', ns)
            nfe_data['VL_ICMSST'] = v_st.text if v_st is not None else "0"
            
            # vFCPST é usado para FECP em notas com ST
            v_fcpst = icms_tot.find('nfe:vFCPST', ns)
            if v_fcpst is not None:
                nfe_data['VL_FECP'] = v_fcpst.text
                
            v_fcp_uf = icms_tot.find('nfe:vFCPUFDest', ns)
            nfe_data['vFCPUFDest'] = v_fcp_uf.text if v_fcp_uf is not None else "0"
        else:
            nfe_data['VL_ICMS_UF_DEST'] = "0"
            nfe_data['VL_FECP'] = "0"
            nfe_data['VL_ICMSST'] = "0"
            nfe_data['vFCPUFDest'] = "0"
           
        # Extraindo COD_PART de infAdic ou infCpl (JÁ ATUALIZADO NO PASSO ANTERIOR)
        nfe_data['COD_PART'] = ""
        texto_total = ""
        inf_adic = root.find('.//nfe:infAdic', ns)
        if inf_adic is not None:
            inf_cpl = inf_adic.find('nfe:infCpl', ns)
            if inf_cpl is not None and inf_cpl.text:
                texto_total += " " + inf_cpl.text
            inf_fisco = inf_adic.find('nfe:infAdFisco', ns)
            if inf_fisco is not None and inf_fisco.text:
                texto_total += " " + inf_fisco.text

        if texto_total:
            import re
            match = re.search(r'(?:CLIENTE|COD\.?\s?CLIENTE)[:\s]+(\d+)', texto_total, re.IGNORECASE)
            if match:
                nfe_data['COD_PART'] = match.group(1)
            else:
                partes = texto_total.upper().split("CLIENTE")
                if len(partes) > 1:
                    cliente_info = partes[1].strip().split()
                    if cliente_info:
                        match_num = re.search(r'(\d+)', cliente_info[0])
                        nfe_data['COD_PART'] = match_num.group(1) if match_num else cliente_info[0]

        # Extraindo xPed de det/prod e colocando no campo PC_CLIENTE
        nfe_data['PC_CLIENTE'] = ""
        dets = root.findall('.//nfe:det', ns)
        for det in dets:
            prod = det.find('nfe:prod', ns)
            if prod is not None:
                x_ped = prod.find('nfe:xPed', ns)
                if x_ped is not None and x_ped.text:
                    nfe_data['PC_CLIENTE'] = x_ped.text.strip()
                    break

        # Extraindo dados do Emitente
        emit = root.find('.//nfe:emit', ns)
        if emit is not None:
            cnpj_emit = emit.find('nfe:CNPJ', ns)
            nfe_data['CNPJ_EMITENTE'] = cnpj_emit.text if cnpj_emit is not None else ""
            
            nome_emit = emit.find('nfe:xNome', ns)
            nfe_data['RAZÃO_SOCIAL_EMITENTE'] = nome_emit.text if nome_emit is not None else ""
            
            endereco = emit.find('nfe:enderEmit', ns)
            if endereco is not None:
                lgr = endereco.find('nfe:xLgr', ns)
                nro = endereco.find('nfe:nro', ns)
                nfe_data['ENDEREÇO'] = f"{lgr.text if lgr is not None else ''} {nro.text if nro is not None else ''}".strip()
                
                mun_emit = endereco.find('nfe:cMun', ns)
                nfe_data['COD_MUN'] = mun_emit.text[2:] if mun_emit is not None and len(mun_emit.text) > 2 else ""
                
                uf_emit = endereco.find('nfe:UF', ns)
                nfe_data['UF_EMITENTE'] = uf_emit.text if uf_emit is not None else ""
                
                cep_emit = endereco.find('nfe:CEP', ns)
                nfe_data['CEP'] = cep_emit.text if cep_emit is not None else ""
                
                fone_emit = endereco.find('nfe:fone', ns)
                nfe_data['TELEFONE'] = fone_emit.text if fone_emit is not None else ""
            else:
                nfe_data['ENDEREÇO'] = ""
                nfe_data['COD_MUN'] = ""
                nfe_data['UF_EMITENTE'] = ""
                nfe_data['CEP'] = ""
                nfe_data['TELEFONE'] = ""
        else:
            nfe_data['CNPJ_EMITENTE'] = ""
            nfe_data['RAZÃO_SOCIAL_EMITENTE'] = ""
            nfe_data['ENDEREÇO'] = ""
            nfe_data['COD_MUN'] = ""
            nfe_data['UF_EMITENTE'] = ""
            nfe_data['CEP'] = ""
            nfe_data['TELEFONE'] = ""
            
        return nfe_data
    except Exception as e:
        print(f"Erro ao extrair dados do XML {filepath}: {e}")
        return None
# Calculando valores adicionais
def calcular_valores_adicionais(nfe_data):
    nfe_data['VL_EUA'] = round(float(nfe_data['VL_ICMSST']) + float(nfe_data['VL_ICMS_UF_DEST']) + float(nfe_data['VL_FECP'].replace(",", ".")), 2)
    nfe_data['DATA_EUA'] = nfe_data['DT_EMISSÃO']
    nfe_data['MÊS'] = datetime.now().strftime("%m")  # Mês atual em formato MM
    nfe_data['NOME'] = nfe_data['RAZÃO_SOCIAL_TOMADOR']
    if nfe_data['CONTRIBUINTE'] == "9":
        nfe_data['VALOR_TOTAL_GNRE'] = round(float(nfe_data['VL_ICMS_UF_DEST'].replace(",", ".")) + float(nfe_data['vFCPUFDest'].replace(",", ".")) + float(nfe_data['VL_FECP'].replace(",", ".")), 2)
    else:
        nfe_data['VALOR_TOTAL_GNRE'] = round(float(nfe_data['VL_ICMSST']) + float(nfe_data['VL_ICMS_UF_DEST']) + float(nfe_data['vFCPUFDest']) + float(nfe_data['VL_FECP']), 2)
    tipo_documento = {"MG": "10", "SC": "10", "PR": "10", "SE": "10", "BA": "10", "GO": "10", "PI": "24", "RS": "22",
        "MA": "10", "ES": "10", "MT": "10", "PE": "22", "AL": "10", "RJ": "24", "PA": "10", "DF": "10", "MS": "10",
        "GO": "10", "CE" : "10", "MT": "22", "RN": "10", "RO": "10", "AC": "10", "AM": "10", "AP": "10", "PA": "10","PE": "10"}
    
    origem = {
    "AC": nfe_data['CHAVE_NFE'],
    "AL": nfe_data['Nº_NFE'],
    "AM": nfe_data['CHAVE_NFE'],
    "AP": nfe_data['CHAVE_NFE'],
    "BA": nfe_data['Nº_NFE'],
    "CE": nfe_data['Nº_NFE'],
    "DF": nfe_data['Nº_NFE'],
    "ES": nfe_data['Nº_NFE'],
    "GO": nfe_data['Nº_NFE'],
    "MA": nfe_data['Nº_NFE'],
    "MG": nfe_data['Nº_NFE'],
    "MS": nfe_data['Nº_NFE'],
    "MT": nfe_data['CHAVE_NFE'],
    "PA": nfe_data['Nº_NFE'],
    "PB": nfe_data['CHAVE_NFE'],
    "PE": nfe_data['CHAVE_NFE'],
    "PI": nfe_data['CHAVE_NFE'],
    "PR": nfe_data['Nº_NFE'],
    "RJ": nfe_data['CHAVE_NFE'],
    "RN": nfe_data['CHAVE_NFE'],
    "RO": nfe_data['CHAVE_NFE'],
    "RR": nfe_data['CHAVE_NFE'],
    "RS": nfe_data['CHAVE_NFE'],
    "SC": nfe_data['Nº_NFE'],
    "SE": nfe_data['Nº_NFE'],
    "SP": nfe_data['CHAVE_NFE'],
    "TO": nfe_data['CHAVE_NFE']
}

    codigoextra1 ={
        "MG": "74", "SC": "84","PR": "87", "SE": "77", "BA": "84", "GO": "102",  "PI": "0", "RS": "74", "MA": "94", "ES": "0", "MT": "17",
        "PE": "9", "AL": "90", "RJ": "117", "PA": "101", "DF": "65", "MS": "88", "GO": "88","CE": "50", "PE" : "50",}
    campo ={
        "MG": "45",  "SC": "45", "PR": "56", "SE": "45", "BA": "45", "GO": "10", "PI": "0", "RS": "62", "MA": "45", "ES": "0", "MT": "74",
        "PE": "92", "AL": "65", "RJ": "118", "PA": "101", "DF": "65", "MS": "65", "GO": "65","CE": "50","PE" : "50",}
    protocolos = {
        "AL": "104/08", "BA": "104/09","DF": "25/11", "ES": "20/13", "GO": "82/11", "MG": "32/09", "PE": "128/10", "PR": "71/11", "RJ": "32/14",
        "RS": "92/09", "SC": "116/12", "SE": "33/12",}
    protocolosOBS = {
        "AL": "104/08", "BA": "104/09 ", "DF": "25/11", "ES": "20/13", "GO": "82/11", "MG": "32/09", "PE": "128/10", "PR": "71/11",
        "RJ": "32/14", "RS": "92/09", "SC": "116/12", "SE": "33/12",}
    contribuinte = {
        "AL": "100099", "BA": "100099", "DF": "100099", "ES": "100099", "GO": "100099","MG": "100099", "PE": "100099", "PR": "100099", "RJ": "100099",
        "RS": "100099", "SC": "100099", "SE": "100099",}
    naocontribuinte = {
        "AL": "100102", "BA": "100102", "DF": "100102", "ES": "100102", "GO": "100102", "MG": "100102", "PE": "100102", "PR": "100102", "RJ": "100102",
        "RS": "100102", "SC": "100102", "SE": "100102","CE":"100102","MT":"100102", "PA":"100102", "MS":"100102", "AM": "100102", "RN":"100102","RO":"100102","RR": "100102", "AC": "100102", "AP": "100102", "TO": "100102",  "PE": "100102"}
    
    if nfe_data['CONTRIBUINTE'] == "9" and nfe_data.get('IE', '') == "":
        nfe_data['IE'] = "00"

    if nfe_data['CONTRIBUINTE'] == "9":
        uf = nfe_data.get('UF_TOMADOR', '')
        if uf == "PR":
            nfe_data['CAMPO_EXTRA1'] = "107"
        elif uf == "BA":
            nfe_data['CAMPO_EXTRA1'] = "86"
        elif uf == "RJ" or uf == "AM" or uf == "RN" or uf == "RO" or uf == "RR" or uf == "AC" or uf == "AP" or uf == "TO" or uf == "PE":
            nfe_data['CAMPO_EXTRA1'] = "117"
        else:
            nfe_data['CAMPO_EXTRA1'] = codigoextra1.get(uf, "")
        nfe_data['PROTOCOLO_ICMS'] = "NC CONV - 93/15"
        nfe_data['PROTOCOLO_OBS'] = "NC CONV - 93/15"
        nfe_data['CAMPO'] = campo.get(uf, "")
        nfe_data['TIPO'] = tipo_documento.get(uf, "")
        nfe_data['ORIGEM'] = origem.get(uf, "")
        nfe_data['COD_RECEITA'] = naocontribuinte.get(uf, "")
        nfe_data['RENOMEAR'] = f"{nfe_data['Nº_NFE']} - {nfe_data['RAZÃO_SOCIAL_TOMADOR']}-{nfe_data['UF_TOMADOR']} R$ {nfe_data['VALOR_TOTAL_GNRE']}.pdf"
    else:
            uf = nfe_data.get('UF_TOMADOR', '')
            nfe_data['PROTOCOLO_ICMS'] = protocolos.get(uf, "")
            nfe_data['PROTOCOLO_OBS'] = protocolosOBS.get(uf, "")
            nfe_data['CAMPO_EXTRA1'] = codigoextra1.get(uf, "")
            nfe_data['CAMPO'] = campo.get(uf, "")
            nfe_data['TIPO'] = tipo_documento.get(uf, "")
            nfe_data['ORIGEM'] = origem.get(uf, "")
            nfe_data['COD_RECEITA'] = contribuinte.get(uf, "")
            nfe_data['RENOMEAR'] = f"{nfe_data['Nº_NFE']} - {nfe_data['RAZÃO_SOCIAL_TOMADOR']}-{nfe_data['UF_TOMADOR']} R$ {nfe_data['VALOR_TOTAL_GNRE']}.pdf"
    nfe_data['OBS_GNRE'] = f"NF-e {nfe_data['Nº_NFE']} {nfe_data['RAZÃO_SOCIAL_TOMADOR'].split()[0]} CNPJ {nfe_data['CNPJ_TOMADOR']} PROT. {nfe_data['PROTOCOLO_OBS']} PC {nfe_data.get('PC_CLIENTE','')}"
    vl_fecp = float(nfe_data['VL_FECP'].replace(",", ".")) if nfe_data['VL_FECP'] else 0.0
    vl_icmsst = float(nfe_data['VL_ICMSST'].replace(",", ".")) if nfe_data['VL_ICMSST'] else 0.0
    vl_icmsst = float(nfe_data['VL_ICMS_UF_DEST'].replace(",", ".")) if nfe_data['VL_ICMS_UF_DEST'] else 0.0
    nfe_data['VL_FECP_GNRE_EUA'] = str(vl_fecp).replace(",", ".")
    # Associar e-mails para envio de GNREs

    
        
    return nfe_data

def atualizar_emails_por_cod_part():
    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()

        cursor.execute("SELECT COD_PART FROM DADOS_GNRE")
        codigos = cursor.fetchall()

        for (cod,) in codigos:
            email = buscar_email_por_cod(cod)
            if email:
                cursor.execute("UPDATE DADOS_GNRE SET EMAIL = ? WHERE COD_PART = ?", (email, cod))

        conn.commit()
        messagebox.showinfo("Sucesso", "E-mails atualizados com sucesso com base no banco.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao atualizar e-mails: {e}")
    finally:
        conn.close()


def atualizar_email(campo_identificador, valor_identificador, novo_email):
    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()

        # Atualiza o campo EMAIL onde o identificador (ex: Nº_NFE ou CNPJ_TOMADOR) for igual ao valor fornecido
        cursor.execute(f'''
            UPDATE DADOS_GNRE
            SET EMAIL = ?
            WHERE {campo_identificador} = ?
        ''', (novo_email, valor_identificador))

        conn.commit()

        if cursor.rowcount == 0:
            messagebox.showinfo("Aviso", "Nenhum registro encontrado para atualizar.")
        else:
            messagebox.showinfo("Sucesso", "E-mail atualizado com sucesso.")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao atualizar e-mail: {e}")
    finally:
        conn.close()
print (atualizar_email)
def inserir_dados(nfe_data):
    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()
    try:
        # Verifica se já existe um registro com a mesma CHAVE_NFE no banco de dados
        cursor.execute("SELECT COUNT(*) FROM DADOS_GNRE WHERE CHAVE_NFE = ?", (nfe_data['CHAVE_NFE'],))
        if cursor.fetchone()[0] == 0:
            # Verificar e converter os valores para float
            vFCPUFDEST = float(nfe_data['vFCPUFDest'].replace(",", ".")) if nfe_data['vFCPUFDest'] else 0.0
            vl_fecp = float(nfe_data['VL_FECP'].replace(",", ".")) if nfe_data['VL_FECP'] else 0.0
            vl_icmsst = float(nfe_data['VL_ICMSST'].replace(",", ".")) if nfe_data['VL_ICMSST'] else 0.0
            vl_icms_uf_dest = float(nfe_data['VL_ICMS_UF_DEST'].replace(",", ".")) if nfe_data['VL_ICMS_UF_DEST'] else 0.0
            # Calcular e atribuir o valor total GNRE
            nfe_data['VALOR_TOTAL_GNRE'] = vl_fecp + vl_icmsst + vl_icms_uf_dest + vFCPUFDEST
            # Exibir valores para depuração
            print(f"VL_FECP: {vl_fecp}, VL_ICMSST: {vl_icmsst}, vFCPUFDEST: {vFCPUFDEST} ,VL_ICMS_UF_DEST: {vl_icms_uf_dest}")
            print(f"VALOR_TOTAL_GNRE: {nfe_data['VALOR_TOTAL_GNRE']}")  # Exibe o total calculado
            # Inserir dados no banco de dados
            cursor.execute('''
                INSERT INTO DADOS_GNRE (Nº_NFE, DT_EMISSÃO, COD_PART, RAZÃO_SOCIAL_TOMADOR, UF_TOMADOR, CONTRIBUINTE, VL_ICMS_UF_DEST, VL_FECP, VL_ICMSST, CHAVE_NFE, 
                    VL_EUA, DATA_EUA, PROTOCOLO_ICMS, PROTOCOLO_OBS, CNPJ_TOMADOR, IE, MÊS, NOME, OBS_GNRE, VL_FECP_GNRE_EUA, PC_CLIENTE, CAMPO_EXTRA1, CAMPO, CNPJ_EMITENTE, RAZÃO_SOCIAL_EMITENTE, 
                    ENDEREÇO, COD_MUN, UF_EMITENTE, CEP, TELEFONE, TIPO, ORIGEM, COD_RECEITA, VALOR_TOTAL_GNRE, RENOMEAR, vFCPUFDest, MUNICIPIO, CAMINHO_PDF)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?,?,?,?,?,?,?)
            ''', (
                nfe_data['Nº_NFE'], nfe_data['DT_EMISSÃO'], nfe_data['COD_PART'], nfe_data['RAZÃO_SOCIAL_TOMADOR'], nfe_data['UF_TOMADOR'],
                nfe_data['CONTRIBUINTE'], nfe_data['VL_ICMS_UF_DEST'], nfe_data['VL_FECP'],  nfe_data['VL_ICMSST'], nfe_data['CHAVE_NFE'],
                nfe_data['VL_EUA'], nfe_data['DATA_EUA'], nfe_data['PROTOCOLO_ICMS'], nfe_data['PROTOCOLO_OBS'], nfe_data['CNPJ_TOMADOR'],
                nfe_data['IE'], nfe_data['MÊS'], nfe_data['NOME'], nfe_data['OBS_GNRE'], nfe_data['VL_FECP_GNRE_EUA'], nfe_data.get('PC_CLIENTE', ''),
                nfe_data['CAMPO_EXTRA1'], nfe_data['CAMPO'], nfe_data['CNPJ_EMITENTE'], nfe_data['RAZÃO_SOCIAL_EMITENTE'], nfe_data['ENDEREÇO'],
                nfe_data['COD_MUN'], nfe_data['UF_EMITENTE'], nfe_data['CEP'], nfe_data['TELEFONE'],  nfe_data['TIPO'], nfe_data['ORIGEM'],
                nfe_data['COD_RECEITA'], "{:.2f}".format(nfe_data['VALOR_TOTAL_GNRE']), nfe_data['RENOMEAR'], nfe_data['vFCPUFDest'], nfe_data['MUNICIPIO'],
                nfe_data.get('CAMINHO_PDF', '')
            ))
            conn.commit()
        else:
            print(f"Arquivo com chave NFE {nfe_data['CHAVE_NFE']} já existe no banco de dados.")
    except sqlite3.Error as e:
        print(f"Erro ao inserir dados: {e}")
    finally:
        conn.close()
def somar_campos():
    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()
    cursor.execute('''
        SELECT SUM(CAST(VL_FECP AS REAL) + CAST(VL_ICMSST AS REAL) + CAST(VL_ICMS_UF_DEST AS REAL)) FROM DADOS_GNRE
    ''')
    resultado = cursor.fetchone()
    total = resultado[0] if resultado[0] is not None else 0.0  # Usar 0.0 se não houver resultado
    conn.close()
    return total
# Função para importar arquivos XML
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def importar_xmls():
    """
    Importa todos os arquivos XML de NF-e encontrados na pasta selecionada.
    Agora aceita TODAS as notas, não apenas as que têm ST.
    Marca quais notas têm valores de ST para facilitar o controle.
    """
    # Seleciona a pasta raiz onde estarão os XMLs
    pasta_raiz = filedialog.askdirectory(title="Selecione a pasta com os arquivos XML")
    if not pasta_raiz:
        return

    # Percorre pastas e subpastas para coletar os arquivos XML
    filepaths = []
    for root, _, files in os.walk(pasta_raiz):
        for file in files:
            if file.endswith(".xml"):
                filepaths.append(os.path.join(root, file))
                
    if not filepaths:
        messagebox.showinfo("Informação", "Nenhum arquivo XML encontrado na pasta selecionada.")
        return

    # Contadores para estatísticas
    total_importados = 0
    total_com_st = 0
    total_sem_st = 0
    total_erros = 0
    
    for i, filepath in enumerate(filepaths):
        try:
            nfe_data = extrair_dados_xml(filepath)
            
            # Pula se o arquivo não for uma NF-e válida (ex: CC-e ou erro na leitura)
            if nfe_data is None:
                continue
                
            # Pula se não encontrar o código do participante (cliente) nas observações
            if not nfe_data.get('COD_PART'):
                print(f"Ignorado: {filepath} - COD_PART não encontrado.")
                continue
                
            # Verificar se tem valores de ST (ICMS, FECP ou ICMSST)
            tem_st = (float(nfe_data.get('VL_ICMS_UF_DEST', 0)) > 0 or 
                     float(nfe_data.get('VL_FECP', 0)) > 0 or 
                     float(nfe_data.get('VL_ICMSST', 0)) > 0)
            
            # Calcular valores adicionais (sempre, para todas as notas)
            nfe_data = calcular_valores_adicionais(nfe_data)
            
            # Marcar notas SEM ST na coluna CAMINHO_PDF para não aparecerem na geração de guias
            if not tem_st:
                nfe_data['CAMINHO_PDF'] = 'SEM ST'
            
            # Inserir no banco de dados (TODAS as notas, não apenas as com ST)
            inserir_dados(nfe_data)
            
            total_importados += 1
            if tem_st:
                total_com_st += 1
            else:
                total_sem_st += 1
            
            app.update_idletasks()
            
        except Exception as e:
            total_erros += 1
            print(f"Erro ao processar {filepath}: {e}")
            # Continua processando os demais arquivos

    # Mensagem final com estatísticas
    mensagem = f"""Importação Concluída!

📊 Estatísticas:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ Total importado: {total_importados} notas
📋 Com ST/ICMS: {total_com_st} notas
📄 Sem ST/ICMS: {total_sem_st} notas
❌ Erros: {total_erros} arquivos

Agora você pode usar todas as notas importadas
para enviar XMLs aos clientes!"""
    
    messagebox.showinfo("Sucesso", mensagem)
    
    # Vincular e-mails automaticamente após a importação
    try:
        atualizar_emails_por_cod_part()
    except Exception as e:
        print(f"Erro ao vincular e-mails automaticamente: {e}")

    # Atualizar tabelas
    try:
        atualizar_todas_as_tabelas_e_abas()
    except:
        pass
    
    
    

# Exemplo de função para atualizar a tabela na interface Tkinter (se necessário)
# Função atualizar_tabela_gnre removida (duplicada e obsoleta)

def executar_monitoramento_xml():
    """Gerencia o timer e dispara a verificação em segundo plano."""
    try:
        # Pega a referência global e verifica de forma segura se o Tkinter ainda está ativo
        _app = globals().get('app')
        if _app is None: return
        # Garante que não haverá erro 'invalid command name' ao fechar
        try:
            if not _app.winfo_exists(): return
        except (tk.TclError, Exception): return
    except (NameError, AttributeError, Exception):
        return
        
    if not hasattr(app, 'monitoramento_ativo'):
        return

    # Se estiver desativado, reseta o tempo e apenas reagenda
    if not app.monitoramento_ativo.get():
        conf = get_configuracoes()
        try: app.tempo_restante_monitor = int(conf.get("MONITOR_INTERVALO", 30))
        except: app.tempo_restante_monitor = 30
        lbl_timer_monitor.config(text="")
        try:
            if app.winfo_exists():
                app.after(1000, executar_monitoramento_xml)
        except:
            pass
        return

    # Atualiza o countdown
    if not hasattr(app, 'tempo_restante_monitor'):
        app.tempo_restante_monitor = 30
    
    app.tempo_restante_monitor -= 1
    
    # Atualiza a label do timer
    lbl_timer_monitor.config(text=f"({app.tempo_restante_monitor}s)")

    if app.tempo_restante_monitor > 0:
        try:
            if app.winfo_exists():
                app.after(1000, executar_monitoramento_xml)
        except:
            pass
        return

    # Se chegou a zero, inicia o processamento em BACKGROUND (thread)
    conf = get_configuracoes()
    try: app.tempo_restante_monitor = int(conf.get("MONITOR_INTERVALO", 30))
    except: app.tempo_restante_monitor = 30
    
    if not app.winfo_exists(): return 
    lbl_timer_monitor.config(text="(Lendo...)")
    
    # Dispara a thread para não travar o sistema
    threading.Thread(target=tarefa_background_xml, daemon=True).start()
    
    # Reagenda o timer para daqui a 1 segundo
    try:
        _app = globals().get('app')
        if _app is not None:
            try:
                if _app.winfo_exists():
                    _app.after(1000, executar_monitoramento_xml)
            except (tk.TclError, Exception): pass
    except (NameError, AttributeError, Exception):
        pass

def tarefa_background_xml():
    """Lógica de importação que roda em segundo plano (Thread)."""
    try:
        # 0. Determina a pasta de monitoramento DINAMICAMENTE
        config = get_configuracoes()
        pasta_base = config.get('PASTA_XML_NFE_ROOT', PASTA_XML_NFE_PADRAO)
        
        # Mapeamento de meses para garantir o nome correto independente de locale
        mapa_meses = {
            "01": "Janeiro", "02": "Fevereiro", "03": "Março", "04": "Abril",
            "05": "Maio", "06": "Junho", "07": "Julho", "08": "Agosto",
            "09": "Setembro", "10": "Outubro", "11": "Novembro", "12": "Dezembro"
        }
        
        agora = datetime.now()
        ano = agora.strftime("%Y")
        mes_num = agora.strftime("%m")
        mes_nome = mapa_meses.get(mes_num, "Janeiro")
        
        # Constrói o caminho: S:\NFE\2026\02 - Fevereiro
        pasta_monitor = os.path.join(pasta_base, ano, f"{mes_num} - {mes_nome}")
        
        if not os.path.exists(pasta_monitor):
            # Se a pasta do mês não existir, não faz nada (ou cria se preferir, mas fiscal costuma já ter)
            return
            
        arquivos = [f for f in os.listdir(pasta_monitor) if f.lower().endswith(".xml")]
        
        if not arquivos:
            return

        # 1. Carrega todas as chaves existentes de uma vez para comparação rápida na memória
        # Isso é MUITO mais rápido do que fazer um SELECT para cada arquivo
        try:
            conn_list = sqlite3.connect("DADOS_GNRE.db")
            cursor_list = conn_list.cursor()
            cursor_list.execute("SELECT CHAVE_NFE FROM DADOS_GNRE")
            chaves_existentes = {row[0] for row in cursor_list.fetchall()}
            conn_list.close()
        except Exception as e:
            print(f"Erro ao carregar chaves do banco: {e}")
            return

        novas_notas = 0
        for arquivo in arquivos:
            caminho_completo = os.path.join(pasta_monitor, arquivo)
            
            try:
                # 2. Extrai apenas a chave para um check rápido
                tree_xml = ET.parse(caminho_completo)
                root_xml = tree_xml.getroot()
                ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
                inf_nfe = root_xml.find('.//nfe:infNFe', ns)
                
                if inf_nfe is not None:
                    chave = inf_nfe.attrib.get('Id', '')[3:]
                    
                    # 3. Verifica se a chave já está no nosso set da memória
                    if chave not in chaves_existentes:
                        # 4. Só agora faz o processamento completo e pesado
                        nfe_data = extrair_dados_xml(caminho_completo)
                        if nfe_data:
                            # 4. Sincroniza com a lógica manual: Pula se não tiver COD_PART
                            if not nfe_data.get('COD_PART'):
                                chaves_existentes.add(chave) # Para não tentar de novo nesta rodada
                                continue

                            nfe_data = calcular_valores_adicionais(nfe_data)
                            
                            tem_st = (float(nfe_data.get('VL_ICMS_UF_DEST', 0)) > 0 or 
                                     float(nfe_data.get('VL_FECP', 0)) > 0 or 
                                     float(nfe_data.get('VL_ICMSST', 0)) > 0)
                            
                            if not tem_st:
                                nfe_data['CAMINHO_PDF'] = 'SEM ST'
                            
                            # Salva no banco
                            inserir_dados(nfe_data)
                            atualizar_emails_por_cod_part_silencioso()
                            
                            novas_notas += 1
                            chaves_existentes.add(chave) # Evita duplicatas na mesma rodada
                            
                            # ATUALIZAÇÃO IMEDIATA: Manda a UI atualizar para cada nota encontrada
                            try:
                                app.after(0, atualizar_todas_as_tabelas_e_abas)
                            except:
                                pass # App pode estar fechando
            except:
                continue # Pula arquivos com erro (ex: temporários abertos)
        
        if novas_notas > 0:
            print(f"Auto-importação finalizada: {novas_notas} novas notas adicionadas ao banco.")
            
    except Exception as e:
        print(f"Erro na thread de monitoramento: {e}")


def executar_organizador_xml():
    """Gerencia o timer do robô organizador de XML."""
    try:
        _app = globals().get("app")
        if _app is None or not _app.winfo_exists(): return
        
        if not hasattr(_app, "organizador_ativo") or not _app.organizador_ativo.get():
            _app.after(10000, executar_organizador_xml)
            return

        if not hasattr(_app, "tempo_organizador"): 
            conf = get_configuracoes()
            try: _app.tempo_organizador = int(conf.get("MONITOR_INTERVALO", 30)) * 2 # Opcional: dobro do tempo
            except: _app.tempo_organizador = 60
        
        _app.tempo_organizador -= 1
        if _app.tempo_organizador <= 0:
            conf = get_configuracoes()
            try: _app.tempo_organizador = int(conf.get("MONITOR_INTERVALO", 30)) * 2
            except: _app.tempo_organizador = 60
            import threading
            threading.Thread(target=tarefa_background_organizador_xml, daemon=True).start()
        
        _app.after(1000, executar_organizador_xml)
    except: pass

def tarefa_background_organizador_xml():
    r"""Lê XMLs da pasta fonte, extrai dhEmi e move para S:\NFE\Ano\Mes."""
    try:
        conf = get_configuracoes()
        # Garante caminhos válidos mesmo que o banco retorne None
        fonte = conf.get("PASTA_FONTE_ORGANIZADOR") or r"C:\Users\Wesley.Raimundo\Desktop\Xml"
        destino_base = conf.get("PASTA_XML_NFE_ROOT") or r"S:\NFE"
        
        if not fonte or not os.path.exists(fonte): return
        
        mapa_meses = {
            "01": "Janeiro", "02": "Fevereiro", "03": "Março", "04": "Abril",
            "05": "Maio", "06": "Junho", "07": "Julho", "08": "Agosto",
            "09": "Setembro", "10": "Outubro", "11": "Novembro", "12": "Dezembro"
        }
        
        arquivos = [f for f in os.listdir(fonte) if f.lower().endswith(".xml")]
        processados = 0
        
        for arquivo in arquivos:
            caminho_xml = os.path.join(fonte, arquivo)
            try:
                import xml.etree.ElementTree as ET
                tree = ET.parse(caminho_xml)
                root = tree.getroot()
                ns = {"nfe": "http://www.portalfiscal.inf.br/nfe"}
                dh_emi = root.find(".//nfe:dhEmi", ns)
                if dh_emi is not None:
                    data_str = dh_emi.text
                    ano = data_str[:4]
                    mes_num = data_str[5:7]
                    mes_nome = mapa_meses.get(mes_num, "Janeiro")
                    
                    pasta_ano = os.path.join(destino_base, ano)
                    pasta_mes = os.path.join(pasta_ano, f"{mes_num} - {mes_nome}")
                    
                    if not os.path.exists(pasta_mes): os.makedirs(pasta_mes, exist_ok=True)
                    
                    # Mover o arquivo
                    import shutil
                    dest_final = os.path.join(pasta_mes, arquivo)
                    if os.path.exists(dest_final): os.remove(dest_final)
                    shutil.move(caminho_xml, dest_final)
                    processados += 1
            except: continue
        
        if processados > 0:
            print(f"Robô Organizador: {processados} XMLs movidos para a rede.")
    except Exception as e: print(f"Erro no Organizar: {e}")

def atualizar_tabela2():
    for row in tree.get_children():
        tree.delete(row)
    
    def fmt(v):
        try:
            val = float(str(v).replace(".", "").replace(",", ".")) if v else 0.0
            return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except: return "R$ 0,00"

    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM DADOS_GNRE WHERE (COD_PART IS NOT NULL AND COD_PART != '') ORDER BY DT_EMISSÃO DESC, Nº_NFE DESC") 
        linhas = cursor.fetchall()
        
        for row in linhas:
            try:
                cancelada = row[41]
                caminho_pdf = row[34]
                status_ws = row[44] or ""
                protocolo = row[43] or ""
                
                values = (
                    row[0],  # nfe
                    row[1],  # emissao
                    row[4],  # uf
                    row[3],  # razao
                    fmt(row[33]), # total
                    fmt(row[7]),  # fecp
                    fmt(row[8]),  # st
                    status_ws,
                    protocolo
                )
                
                tags = ()
                if cancelada == 1: tags = ("cancelada",)
                elif caminho_pdf: tags = ("anexado",)
                else: tags = ("nao_anexado",)
                
                tree.insert("", "end", values=values, tags=tags)
            except Exception as e: continue
        conn.close()
    except Exception as e:
        print(f"Erro ao carregar tabela: {e}")

def atualizar_aba_consulta_apenas():
    """Atualiza APENAS a aba Consulta (dashboard). Não toca na aba Gerar GNRE."""
    try:
        atualizar_tabela2()
        
        try:
            conn = sqlite3.connect(NOME_BANCO_DADOS)
            cursor = conn.cursor()
            
            # Fórmula robusta para lidar com formatos de número (SQL padrão ou PT-BR com vírgula)
            def sql_soma_col(col):
                return f"SUM(CASE WHEN {col} LIKE '%,%' THEN CAST(REPLACE(REPLACE(IFNULL({col},'0'), '.', ''), ',', '.') AS REAL) ELSE CAST(IFNULL({col},'0') AS REAL) END)"

            # 1. Pendentes Qtd
            cursor.execute("SELECT COUNT(*) FROM DADOS_GNRE WHERE CANCELADA = 0 AND (CAMINHO_PDF IS NULL OR CAMINHO_PDF = '') AND (VL_ICMS_UF_DEST > 0 OR VL_FECP > 0 OR VL_ICMSST > 0)")
            p = cursor.fetchone()[0]
            if "lbl_val_pendentes" in globals(): lbl_val_pendentes.config(text=str(p))

            # 2. Valor Financeiro (Pendentes em R$)
            f_fecp = f"CASE WHEN VL_FECP LIKE '%,%' THEN CAST(REPLACE(REPLACE(IFNULL(VL_FECP,'0'), '.', ''), ',', '.') AS REAL) ELSE CAST(IFNULL(VL_FECP,'0') AS REAL) END"
            f_st = f"CASE WHEN VL_ICMSST LIKE '%,%' THEN CAST(REPLACE(REPLACE(IFNULL(VL_ICMSST,'0'), '.', ''), ',', '.') AS REAL) ELSE CAST(IFNULL(VL_ICMSST,'0') AS REAL) END"
            f_dest = f"CASE WHEN VL_ICMS_UF_DEST LIKE '%,%' THEN CAST(REPLACE(REPLACE(IFNULL(VL_ICMS_UF_DEST,'0'), '.', ''), ',', '.') AS REAL) ELSE CAST(IFNULL(VL_ICMS_UF_DEST,'0') AS REAL) END"
            
            cursor.execute(f"SELECT SUM({f_fecp} + {f_st} + {f_dest}) FROM DADOS_GNRE WHERE CANCELADA = 0 AND (CAMINHO_PDF IS NULL OR CAMINHO_PDF = '')")
            f_pend = cursor.fetchone()[0] or 0
            if "lbl_val_financeiro" in globals(): lbl_val_financeiro.config(text=f"R$ {f_pend:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            # 3. Histórico Total (Qtd)
            cursor.execute("SELECT COUNT(*) FROM ENVIADOS_GNRE")
            h = cursor.fetchone()[0]
            if "lbl_val_hoje" in globals(): lbl_val_hoje.config(text=str(h))

            # 4. Valor Total Gerado (R$) - Soma de FECP + ICMSST (SEM FILTROS, conforme solicitado pelo valor alvo)
            cursor.execute(f"SELECT SUM({f_fecp} + {f_st}) FROM DADOS_GNRE")
            tg = cursor.fetchone()[0] or 0
            if "lbl_val_total_geral" in globals(): lbl_val_total_geral.config(text=f"R$ {tg:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            # 5. Alertas
            d2 = (datetime.now() - timedelta(days=2)).strftime("%Y-%m-%d")
            cursor.execute("SELECT COUNT(*) FROM DADOS_GNRE WHERE CANCELADA = 0 AND (CAMINHO_PDF IS NULL OR CAMINHO_PDF = '') AND DT_EMISSÃO < ?", (d2,))
            a = cursor.fetchone()[0]
            if "lbl_val_alertas" in globals(): lbl_val_alertas.config(text=str(a))

            # 6. Status Robôs
            if "app" in globals():
                s_imp = app.monitoramento_ativo.get()
                s_org = (app.organizador_ativo.get() if hasattr(app, "organizador_ativo") else False)
                
                if "lbl_dash_imp_status" in globals():
                    lbl_dash_imp_status.config(text="ONLINE 🟢" if s_imp else "OFFLINE ⚪", fg=COR_SUCESSO if s_imp else "#95a5a6")
                if "lbl_dash_org_status" in globals():
                    lbl_dash_org_status.config(text="ONLINE 🟢" if s_org else "OFFLINE ⚪", fg=COR_SUCESSO if s_org else "#95a5a6")

            conn.close()
        except Exception as e: print(f"Erro Dash: {e}")
    except Exception as e: print(f"Erro Tabelas: {e}")

def atualizar_todas_as_tabelas_e_abas():
    """Atualiza TODAS as tabelas e abas (inclui aba Gerar GNRE). Chamada apenas manualmente."""
    try:
        atualizar_tabela2()
        if "entrada_nfe" in globals() and not entrada_nfe.get().strip(): listar_nfe_sem_caminho_pdf()
        
        try:
            conn = sqlite3.connect(NOME_BANCO_DADOS)
            cursor = conn.cursor()
            
            # Fórmula robusta para lidar com formatos de número (SQL padrão ou PT-BR com vírgula)
            def sql_soma_col(col):
                return f"SUM(CASE WHEN {col} LIKE '%,%' THEN CAST(REPLACE(REPLACE(IFNULL({col},'0'), '.', ''), ',', '.') AS REAL) ELSE CAST(IFNULL({col},'0') AS REAL) END)"

            # 1. Pendentes Qtd
            cursor.execute("SELECT COUNT(*) FROM DADOS_GNRE WHERE CANCELADA = 0 AND (CAMINHO_PDF IS NULL OR CAMINHO_PDF = '') AND (VL_ICMS_UF_DEST > 0 OR VL_FECP > 0 OR VL_ICMSST > 0)")

            p = cursor.fetchone()[0]
            if "lbl_val_pendentes" in globals(): lbl_val_pendentes.config(text=str(p))

            # 2. Valor Pendente (R$) - ICMS Destino + FCP + ST
            f_fecp = f"CASE WHEN VL_FECP LIKE '%,%' THEN CAST(REPLACE(REPLACE(IFNULL(VL_FECP,'0'), '.', ''), ',', '.') AS REAL) ELSE CAST(IFNULL(VL_FECP,'0') AS REAL) END"
            f_st = f"CASE WHEN VL_ICMSST LIKE '%,%' THEN CAST(REPLACE(REPLACE(IFNULL(VL_ICMSST,'0'), '.', ''), ',', '.') AS REAL) ELSE CAST(IFNULL(VL_ICMSST,'0') AS REAL) END"
            f_dest = f"CASE WHEN VL_ICMS_UF_DEST LIKE '%,%' THEN CAST(REPLACE(REPLACE(IFNULL(VL_ICMS_UF_DEST,'0'), '.', ''), ',', '.') AS REAL) ELSE CAST(IFNULL(VL_ICMS_UF_DEST,'0') AS REAL) END"
            
            cursor.execute(f"SELECT SUM({f_fecp} + {f_st} + {f_dest}) FROM DADOS_GNRE WHERE CANCELADA = 0 AND (CAMINHO_PDF IS NULL OR CAMINHO_PDF = '')")
            f_pend = cursor.fetchone()[0] or 0
            if "lbl_val_financeiro" in globals(): lbl_val_financeiro.config(text=f"R$ {f_pend:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            # 3. Histórico Total (Qtd)
            cursor.execute("SELECT COUNT(*) FROM ENVIADOS_GNRE")
            h = cursor.fetchone()[0]
            if "lbl_val_hoje" in globals(): lbl_val_hoje.config(text=str(h))

            # 4. Valor Total Gerado (R$) - Soma de FECP + ICMSST (SEM FILTROS, conforme solicitado pelo valor alvo)
            cursor.execute(f"SELECT SUM({f_fecp} + {f_st}) FROM DADOS_GNRE")
            tg = cursor.fetchone()[0] or 0
            if "lbl_val_total_geral" in globals(): lbl_val_total_geral.config(text=f"R$ {tg:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            # 5. Alertas
            d2 = (datetime.now() - timedelta(days=2)).strftime("%Y-%m-%d")
            cursor.execute("SELECT COUNT(*) FROM DADOS_GNRE WHERE CANCELADA = 0 AND (CAMINHO_PDF IS NULL OR CAMINHO_PDF = '') AND DT_EMISSÃO < ?", (d2,))
            a = cursor.fetchone()[0]
            if "lbl_val_alertas" in globals(): lbl_val_alertas.config(text=str(a))

            # 6. Status Robôs
            if "app" in globals():
                s_imp = app.monitoramento_ativo.get()
                s_org = (app.organizador_ativo.get() if hasattr(app, "organizador_ativo") else False)
                
                if "lbl_dash_imp_status" in globals():
                    lbl_dash_imp_status.config(text="ONLINE 🟢" if s_imp else "OFFLINE ⚪", fg=COR_SUCESSO if s_imp else "#95a5a6")
                if "lbl_dash_org_status" in globals():
                    lbl_dash_org_status.config(text="ONLINE 🟢" if s_org else "OFFLINE ⚪", fg=COR_SUCESSO if s_org else "#95a5a6")

            conn.close()
        except Exception as e: print(f"Erro Dash: {e}")
    except Exception as e: print(f"Erro Tabelas: {e}")

def criar_aba_dashboard(parent):
    """Cria a aba de dashboard com estatísticas visuais e financeiras."""
    global lbl_val_pendentes, lbl_val_hoje, lbl_val_alertas, lbl_val_financeiro, lbl_val_robo, lbl_val_total_geral
    
    aba = tk.Frame(parent, bg="#f5f6fa")
    container = tk.Frame(aba, bg="#f5f6fa", padx=40, pady=30)
    container.pack(fill="both", expand=True)
    
    header = tk.Frame(container, bg="#f5f6fa")
    header.pack(fill="x", pady=(0, 20))
    tk.Label(header, text="📊 DASHBOARD ESTRATÉGICO GNRE", font=("Segoe UI", 22, "bold"), bg="#f5f6fa", fg="#2c3e50").pack(side="left")
    tk.Button(header, text="🔄 Atualizar Painel", font=("Segoe UI", 9, "bold"), bg="#3498db", fg="white", 
              padx=20, pady=10, relief="flat", cursor="hand2", command=atualizar_todas_as_tabelas_e_abas).pack(side="right")

    grid_frame = tk.Frame(container, bg="#f5f6fa")
    grid_frame.pack(fill="both", expand=True)

    def criar_card(m, tit, val, cor, row, col):
        c = tk.Frame(m, bg="white", highlightthickness=1, highlightbackground="#dcdde1", padx=20, pady=20)
        c.grid(row=row, column=col, sticky="nsew", padx=10, pady=10)
        tk.Label(c, text=tit, font=("Segoe UI", 9, "bold"), bg="white", fg="#7f8c8d").pack(anchor="w")
        l = tk.Label(c, text=val, font=("Segoe UI", 24, "bold"), bg="white", fg=cor)
        l.pack(anchor="w", pady=(5, 0))
        tk.Frame(c, bg=cor, height=4).place(relx=0, rely=0, relwidth=1)
        return l

    grid_frame.columnconfigure((0,1,2), weight=1)
    
    lbl_val_pendentes = criar_card(grid_frame, "NOTAS PENDENTES", "0", "#f39c12", 0, 0)
    lbl_val_hoje = criar_card(grid_frame, "HISTÓRICO TOTAL (Qtd)", "0", "#27ae60", 0, 1)
    lbl_val_alertas = criar_card(grid_frame, "ALERTAS (VENCIDOS)", "0", "#e74c3c", 0, 2)
    
    lbl_val_financeiro = criar_card(grid_frame, "VALOR PENDENTE (R$)", "R$ 0,00", "#d35400", 1, 0)
    lbl_val_total_geral = criar_card(grid_frame, "VALOR TOTAL GERADO (R$)", "R$ 0,00", "#2c3e50", 1, 1)
    
    # 6º Card: Sessão / Informativo
    def criar_card_info(m, tit, val, cor, row, col):
        c = tk.Frame(m, bg="white", highlightthickness=1, highlightbackground="#dcdde1", padx=20, pady=20)
        c.grid(row=row, column=col, sticky="nsew", padx=10, pady=10)
        tk.Label(c, text=tit, font=("Segoe UI", 9, "bold"), bg="white", fg="#7f8c8d").pack(anchor="w")
        tk.Label(c, text=val, font=("Segoe UI", 14, "bold"), bg="white", fg=COR_SECUNDARIA).pack(anchor="w", pady=(10, 0))
        tk.Frame(c, bg=cor, height=4).place(relx=0, rely=0, relwidth=1)
        return c

    criar_card_info(grid_frame, "SISTEMA OPERACIONAL", f"{platform.system()} {platform.release()}", "#95a5a6", 1, 2)

    # ================= CPU / CENTRAL DE AUTOMAÇÃO (ROBÔS INTEGRADOS) =================
    extra_f = tk.Frame(container, bg="white", highlightthickness=1, highlightbackground="#dcdde1", padx=25, pady=20)
    extra_f.pack(fill="x", pady=20)
    tk.Frame(extra_f, bg=COR_DESTAQUE, width=6).pack(side="left", fill="y", padx=(0, 20))
    
    # Lado Esquerdo: Texto de Monitoramento
    info_col = tk.Frame(extra_f, bg="white")
    info_col.pack(side="left", fill="both", expand=True)
    
    tk.Label(info_col, text="🚀 CENTRAL DE MONITORAMENTO E AUTOMAÇÃO", font=("Segoe UI", 11, "bold"), bg="white", fg=COR_PRIMARIA).pack(anchor="w")
    tk.Label(info_col, text="O sistema monitora a rede e o desktop em tempo real. Ciclo de atualização: 15 segundos.", 
             font=("Segoe UI", 9), bg="white", fg="#7f8c8d").pack(anchor="w", pady=(2, 0))
    
    # Lado Direito: Robôs com Status Visual Moderno
    robo_col = tk.Frame(extra_f, bg="white")
    robo_col.pack(side="right", padx=10)
    
    def create_status_badge(parent, icon, title):
        f = tk.Frame(parent, bg="#f8f9fa", padx=15, pady=8, highlightthickness=1, highlightbackground="#dcdde1")
        f.pack(side="left", padx=5)
        tk.Label(f, text=f"{icon} {title}", font=("Segoe UI", 9, "bold"), bg="#f8f9fa", fg=COR_SECUNDARIA).pack(side="left", padx=(0, 10))
        lbl = tk.Label(f, text="OFFLINE", font=("Segoe UI", 9, "bold"), bg="#f8f9fa", fg="#95a5a6")
        lbl.pack(side="left")
        return lbl

    global lbl_dash_imp_status, lbl_dash_org_status
    lbl_dash_imp_status = create_status_badge(robo_col, "📥", "IMPORTADOR")
    lbl_dash_org_status = create_status_badge(robo_col, "📂", "ORGANIZADOR")

    def auto_refresh():
        if parent.winfo_exists():
            atualizar_aba_consulta_apenas(); parent.after(15000, auto_refresh)
    parent.after(3000, auto_refresh)
    return aba

def importar_emails_por_arquivo():
    caminho_arquivo = filedialog.askopenfilename(title="Selecione o arquivo .txt de e-mails", filetypes=[("Arquivos de texto", "*.txt")])
    if not caminho_arquivo:
        return

    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()

        with open(caminho_arquivo, "r", encoding="utf-8") as f:
            for linha in f:
                partes = linha.strip().split("|")
                if len(partes) == 2:
                    cod_part, email = partes[0].strip(), partes[1].strip()
                    cursor.execute("""
                        INSERT INTO EMAIL_CLIENTES (COD_PART, EMAIL)
                        VALUES (?, ?)
                        ON CONFLICT(COD_PART) DO UPDATE SET EMAIL=excluded.EMAIL
                    """, (cod_part, email))
        
        conn.commit()
        messagebox.showinfo("Sucesso", "E-mails importados com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao importar e-mails: {e}")
    finally:
        conn.close()

# ========== NOVOS MÓDULOS INTEGRADOS ==========

def organizar_comprovantes():
    """Lógica do script Copiar Comprovante.py integrada."""
    config = get_configuracoes()
    origem = r"T:\GA\GNRE - COMPROVANTES"
    destino_base = config.get('PASTA_GNRE_ROOT', PASTA_GNRE_PADRAO)
    
    if not os.path.exists(origem):
        messagebox.showerror("Erro", f"Pasta de origem não encontrada:\n{origem}")
        return

    meses_pt = {
        1: "01 - Janeiro", 2: "02 - Fevereiro", 3: "03 - Março", 4: "04 - Abril",
        5: "05 - Maio", 6: "06 - Junho", 7: "07 - Julho", 8: "08 - Agosto",
        9: "09 - Setembro", 10: "10 - Outubro", 11: "11 - Novembro", 12: "12 - Dezembro"
    }

    movidos = 0
    erros = 0
    
    for raiz, _, arquivos in os.walk(origem):
        for arquivo in arquivos:
            if not arquivo.lower().endswith(".pdf"):
                continue

            caminho_arquivo = os.path.join(raiz, arquivo)
            try:
                # Usa data de modificação ou criação
                timestamp = os.path.getmtime(caminho_arquivo)
                data = datetime.fromtimestamp(timestamp)
                ano = data.year
                mes = data.month
                pasta_mes = meses_pt.get(mes, f"{mes:02d}")

                destino_final = os.path.join(destino_base, str(ano), pasta_mes, "COMPROVANTE")
                os.makedirs(destino_final, exist_ok=True)
                
                shutil.move(caminho_arquivo, os.path.join(destino_final, arquivo))
                movidos += 1
            except Exception as e:
                print(f"Erro ao mover {arquivo}: {e}")
                erros += 1

    messagebox.showinfo("Concluído", f"Organização finalizada!\n\nMovidos: {movidos}\nErros: {erros}")

def split_pdf_internal(pdf_path, output_dir):
    """Divide PDF em páginas individuais para renomeação."""
    try:
        pdf = PdfReader(pdf_path)
        file_list = []
        for page_num in range(len(pdf.pages)):
            pdf_writer = PdfWriter()
            pdf_writer.add_page(pdf.pages[page_num])
            out_path = os.path.join(output_dir, f"pagina_{page_num + 1}.pdf")
            with open(out_path, 'wb') as out:
                pdf_writer.write(out)
            file_list.append(out_path)
        return file_list
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao dividir PDF: {e}")
        return []

def renomear_lote_pdf_gui():
    """Lógica do script Renomar GNRE.py integrada."""
    try:
        conn = sqlite3.connect(NOME_BANCO_DADOS)
        cursor = conn.cursor()
        cursor.execute("SELECT renomear, VALOR_TOTAL_GNRE FROM DADOS_GNRE WHERE (CAMINHO_PDF IS NULL OR CAMINHO_PDF = '') AND CANCELADA = 0 ORDER BY RENOMEAR")
        rename_data = cursor.fetchall()
        conn.close()
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler banco: {e}")
        return

    if not rename_data:
        messagebox.showinfo("Aviso", "Nenhum item pendente de renomeação no banco.")
        return

    rename_dict = {nome: valor for nome, valor in rename_data}
    pdf_path = filedialog.askopenfilename(title="Selecione o PDF com o LOTE de guias", filetypes=[("PDF files", "*.pdf")])
    if not pdf_path: return

    output_dir = os.path.join(PASTA_BACKUP_PADRAO, datetime.now().strftime("%Y"), datetime.now().strftime("%Y-%m-%d"))
    os.makedirs(output_dir, exist_ok=True)

    file_list = split_pdf_internal(pdf_path, output_dir)
    if not file_list: return

    selected_names = []
    
    def confirmar_renomeacao_lote():
        nonlocal selected_names
        dialogo = Toplevel(app)
        dialogo.title("Associar Páginas GNRE")
        dialogo.geometry("850x700")
        dialogo.configure(bg="#f8f9fa")
        dialogo.transient(app)
        dialogo.grab_set()
        
        # Centralizar
        dialogo.update_idletasks()
        x = (dialogo.winfo_screenwidth() // 2) - (850 // 2)
        y = (dialogo.winfo_screenheight() // 2) - (700 // 2)
        dialogo.geometry(f"+{int(x)}+{int(y)}")

        # Cabeçalho
        header = tk.Frame(dialogo, bg=COR_PRIMARIA, height=70)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text="🏷️ ASSOCIAÇÃO DE PÁGINAS EM LOTE", font=("Segoe UI", 16, "bold"), bg=COR_PRIMARIA, fg="white").pack(pady=15)

        default_name = "--- SELECIONE O DOCUMENTO ---"
        ignorar_name = "⏩ IGNORAR PÁGINA / PÁGINA ADICIONAL"
        names_options = [default_name, ignorar_name] + sorted(rename_dict.keys())
        
        container = tk.Frame(dialogo, bg="#f8f9fa", padx=30, pady=15)
        container.pack(fill="both", expand=True)

        info_frame = tk.Frame(container, bg="#e3f2fd", padx=15, pady=10, highlightthickness=1, highlightbackground="#90caf9")
        info_frame.pack(fill="x", pady=(0, 15))
        
        tk.Label(info_frame, text=f"📋 Total de páginas no PDF: {len(file_list)}", font=("Segoe UI", 11, "bold"), bg="#e3f2fd", fg="#1976d2").pack(anchor="w")
        tk.Label(info_frame, text="Associe cada página a uma nota fiscal ou use 'IGNORAR PÁGINA' para as páginas que sobrarem.", 
                 font=("Segoe UI", 9), bg="#e3f2fd", fg="#1976d2").pack(anchor="w")

        if len(file_list) > 5:
            tk.Label(container, text="⚠️ Use a barra de rolagem (ou o mouse) para ver todas as páginas abaixo", 
                     font=("Segoe UI", 9, "bold"), bg="#fff3cd", fg="#856404", pady=5).pack(fill="x", pady=(0, 10))

        # Scrollable Area
        outer_frame = tk.Frame(container, bg="white", highlightthickness=1, highlightbackground="#dcdde1")
        outer_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer_frame, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=750)
        
        # Ajustar largura do frame interno quando o canvas redimensionar
        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)
        
        canvas.configure(yscrollcommand=scrollbar.set)

        # Mousewheel support
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        vars_local = []
        for i, f in enumerate(file_list):
            row_f = tk.Frame(scrollable_frame, bg="white", pady=10)
            row_f.pack(fill='x', padx=15)
            
            if i > 0:
                tk.Frame(scrollable_frame, height=1, bg="#f1f2f6").pack(fill='x', padx=25)

            lbl_pag = tk.Label(row_f, text=f"Página {i+1:02d}", font=("Segoe UI", 10, "bold"), bg="white", fg=COR_PRIMARIA, width=10, anchor="w")
            lbl_pag.pack(side='left')
            
            v = tk.StringVar(dialogo)
            v.set(default_name)
            vars_local.append(v)
            
            # Usando tk.OptionMenu para maior compatibilidade com variáveis
            opt = tk.OptionMenu(row_f, v, *names_options)
            opt.config(bg="white", activebackground="#f1f2f6", relief="flat", font=("Segoe UI", 9), anchor="w")
            opt.pack(side='right', fill='x', expand=True, padx=(15, 0))

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        def done():
            # Desvincular mousewheel para não afetar outras telas
            canvas.unbind_all("<MouseWheel>")
            
            res = [v.get().strip() for v in vars_local]
            
            # 1. Verificar se faltou alguma página
            faltando = [i+1 for i, val in enumerate(res) if val == default_name]
            if faltando:
                msg = f"As seguintes páginas não foram associadas:\n\n{', '.join(map(str, faltando))}\n\n"
                msg += "Por favor, selecione uma nota ou 'IGNORAR PÁGINA' para todas as páginas listadas."
                messagebox.showerror("Atenção", msg, parent=dialogo)
                return
            
            # 2. Verificar duplicados (exceto 'ignorar_name')
            real_selections = [r for r in res if r != ignorar_name]
            if len(real_selections) != len(set(real_selections)):
                # Identificar qual está duplicado
                vistos = set()
                duplicados = [x for x in real_selections if x in vistos or vistos.add(x)]
                messagebox.showerror("Erro", f"Você selecionou o mesmo documento para páginas diferentes:\n\n{duplicados[0][:100]}...", parent=dialogo)
                return
            
            # 3. Validar valores se contiverem R$
            for name in real_selections:
                if "R$" in name:
                    try:
                        # Regex melhorada para pegar o valor completo após R$
                        match = re.search(r'R\$\s*([\d.,]+)', name)
                        if match:
                            v_raw = match.group(1)
                            # Lógica robusta para detectar formato do valor (BR ou US)
                            if ',' in v_raw and '.' in v_raw:
                                # Formato 1.234,56 -> 1234.56
                                v_clean = v_raw.replace('.', '').replace(',', '.')
                            elif ',' in v_raw:
                                # Formato 1234,56 -> 1234.56
                                v_clean = v_raw.replace(',', '.')
                            else:
                                # Formato 1234.56 (padrão float Python)
                                v_clean = v_raw
                                
                            v_float = float(v_clean)
                            
                            # Valor do banco (sempre armazenado com . ou vindo formatado)
                            val_banco_str = str(rename_dict[name]).replace(',', '.')
                            val_banco_float = float(val_banco_str)
                            
                            if abs(v_float - val_banco_float) > 0.05: # Tolerância de 5 centavos
                                messagebox.showerror("Valor Divergente", 
                                    f"O valor na nota selecionada ({v_float:.2f}) difere do valor no banco de dados ({val_banco_float:.2f}) para o item:\n\n{name}", 
                                    parent=dialogo)
                                return
                    except Exception as e:
                        print(f"Erro ao validar valor: {e}")
            
            selected_names.extend(res)
            dialogo.destroy()

        footer = tk.Frame(dialogo, bg="#ecf0f1", height=80)
        footer.pack(fill='x', side='bottom')
        footer.pack_propagate(False)

        btn_confirm = tk.Button(footer, text="✅ FINALIZAR E SALVAR ASSOCIAÇÕES", font=("Segoe UI", 11, "bold"), 
                               bg=COR_SUCESSO, fg="white", padx=30, pady=12, relief="flat", command=done, cursor="hand2")
        btn_confirm.pack(pady=15)
        
        return dialogo

    diag = confirmar_renomeacao_lote()
    app.wait_window(diag)

    if not selected_names:
        # Se não salvou nada ou fechou no X, limpar arquivos temporários
        for f in file_list: 
            try: os.remove(f)
            except: pass
        return

    ignorar_name = "⏩ IGNORAR PÁGINA / PÁGINA ADICIONAL"
    
    # Renomeação e atualização no Banco
    erros_renomeacao = []
    sucessos_count = 0
    
    for i, (old, new) in enumerate(zip(file_list, selected_names)):
        try:
            if new == ignorar_name:
                if os.path.exists(old): os.remove(old)
                continue
                
            novo_nome_limpo = re.sub(r'[\\/*?:"<>|]', "", new) # Limpar caracteres inválidos para Windows
            novo_caminho = os.path.join(output_dir, f"{novo_nome_limpo}.pdf")
            
            # Resolver conflito se arquivo já existir
            if os.path.exists(novo_caminho):
                base, ext = os.path.splitext(novo_caminho)
                novo_caminho = f"{base}_{i+1}{ext}"
                
            os.rename(old, novo_caminho)
            
            with sqlite3.connect(NOME_BANCO_DADOS) as conn:
                cursor_upd = conn.execute("UPDATE DADOS_GNRE SET CAMINHO_PDF = ? WHERE RENOMEAR = ?", (novo_caminho, new))
                if cursor_upd.rowcount > 0:
                    sucessos_count += 1
                else:
                    erros_renomeacao.append(f"Registro não encontrado no BD para: {new[:50]}...")
        except Exception as e:
            erros_renomeacao.append(f"Erro ao processar página {i+1} ({new[:30]}): {e}")

    # Limpar PDF original do lote
    try: 
        if os.path.exists(pdf_path): os.remove(pdf_path)
    except: pass
    
    if erros_renomeacao:
        msg_erro = f"Associação concluída com {sucessos_count} sucessos.\n\nPorém, ocorreram os seguintes problemas:\n" + "\n".join(erros_renomeacao[:5])
        if len(erros_renomeacao) > 5: msg_erro += "\n... e outros."
        messagebox.showwarning("Aviso", msg_erro)
    else:
        messagebox.showinfo("Sucesso", f"Lote processado com sucesso!\n\n✅ {sucessos_count} guias foram associadas e salvas.")
        
    atualizar_todas_as_tabelas_e_abas()

def registrar_log_envio(caminho_pdf, email):
    """Registra o envio em um arquivo CSV."""
    try:
        log_file = "log_envio_gnre.csv"
        exists = os.path.isfile(log_file)
        with open(log_file, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if not exists:
                writer.writerow(["Data", "E-mail", "Arquivo"])
            writer.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), email, caminho_pdf])
    except: pass

def enviar_email_com_anexo(tree, selecionado):
    """Envia o email com anexo (PDF e XML) usando o novo template profissional"""
    try:
        dados = tree.item(selecionado, "values")
        # Índices baseados na configuração da Treeview (conforme dica do usuário)
        # Se tree = columns_gnre [0=NFE, 1=UF, 2=VALOR, 3=CLIENTE, 4=OBS] é diferente de Consulta [0=NFE, 1=Data, 2=Cod, 3=Cliente]
        # Mas a função vai buscar no BD se os índices da tree não baterem com o esperado para anexo
        
        nfe_numero = dados[0]
        
        # Como os índices da Treeview variam entre abas, vamos buscar os dados COMPLETOS no BD pelo Nº NFE
        with sqlite3.connect(NOME_BANCO_DADOS) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM DADOS_GNRE WHERE Nº_NFE = ?", (nfe_numero,))
            row = cursor.fetchone()
        
        if not row:
            messagebox.showerror("Erro", "Dados da nota não encontrados no banco.")
            return

        # Mapeamento robusto dos dados do banco
        emissao = row[1]
        chave = row[9]
        cnpj = row[14]
        destinatario = row[3]
        caminho_pdf = row[34]
        caminho_xml = row[36]
        email_destino = row[37]

        if not email_destino:
            messagebox.showwarning("Aviso", "Cliente sem e-mail cadastrado.")
            return

        # Verifica duplicata
        with sqlite3.connect(NOME_BANCO_DADOS) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT 1 FROM ENVIADOS_GNRE WHERE CAMINHO_PDF = ?", (caminho_pdf,))
            if cursor.fetchone():
                if not messagebox.askyesno("Aviso", "Este e-mail já consta como enviado. Deseja reenviar?"):
                    return

        is_especial = caminho_pdf in ["SEM ST", "POR APURAÇÃO"]

        if not is_especial and (not caminho_pdf or not os.path.isfile(caminho_pdf)):
            messagebox.showwarning("Atenção", "Arquivo PDF da guia não encontrado.")
            return

        # Formatar emissão
        if emissao and "-" in emissao:
            try: emissao_fmt = datetime.strptime(emissao, "%Y-%m-%d").strftime("%d/%m/%Y")
            except: emissao_fmt = emissao
        else: emissao_fmt = emissao

        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = email_destino
        
        if is_especial:
            mail.Subject = f"NF-e {nfe_numero}"
            texto_topo = "Segue anexa a NFe abaixo:"
        else:
            mail.Subject = f"CI GNRE - {nfe_numero}"
            texto_topo = "Segue em anexo a guia de recolhimento GNRE referente à nota fiscal abaixo:"

        mail.Body = (
            "Prezado cliente,\n\n"
            f"{texto_topo}\n\n"
            f"NFe: {nfe_numero}\n"
            f"Serie: 3\n"
            f"Emissao: {emissao_fmt}\n"
            f"Chave de Acesso: {chave}\n\n"
            f"Destinatario: {destinatario}\n"
            f"CNPJ: {cnpj}\n\n"
            "Atenciosamente,\n\n"
            "Dinatecnica Industria e Comercio Ltda.\n"
            "Tel.: +11 4785-2230"
        )

        if not is_especial and os.path.isfile(caminho_pdf):
            mail.Attachments.Add(caminho_pdf)
        if caminho_xml and os.path.isfile(caminho_xml):
            mail.Attachments.Add(caminho_xml)

        mail.Send()
        
        # Registrar envio
        registrar_log_envio(caminho_pdf, email_destino)
        with sqlite3.connect(NOME_BANCO_DADOS) as conn:
            conn.execute("INSERT OR IGNORE INTO ENVIADOS_GNRE (CAMINHO_PDF) VALUES (?)", (caminho_pdf,))
        
        messagebox.showinfo("Sucesso", f"E-mail enviado com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro no envio: {e}")

# ========== MÓDULO ASSINAR PDF ==========

def mm_to_points(mm):
    return mm * 2.83465

def assinar_pdf_logic(caminho_pdf_entrada, texto_assinatura, pos_x_mm, pos_y_mm,
                    fonte, tamanho_fonte, pagina_alvo,
                    caminho_imagem=None, img_x_mm=None, img_y_mm=None, img_larg_mm=None, img_alt_mm=None):
    try:
        leitor_pdf = PdfReader(caminho_pdf_entrada)
        escritor_pdf = PdfWriter()
        total_paginas = len(leitor_pdf.pages)

        if pagina_alvo == "Todas":
            paginas_para_assinar = list(range(total_paginas))
        elif pagina_alvo == "Primeira":
            paginas_para_assinar = [0]
        elif pagina_alvo == "Última":
            paginas_para_assinar = [total_paginas - 1]
        else:
            try:
                pagina_index = int(pagina_alvo) - 1
                paginas_para_assinar = [pagina_index] if 0 <= pagina_index < total_paginas else []
            except:
                paginas_para_assinar = [0]

        for i in range(total_paginas):
            pagina = leitor_pdf.pages[i]
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)

            if i in paginas_para_assinar:
                if caminho_imagem and os.path.isfile(caminho_imagem):
                    img_x = mm_to_points(img_x_mm)
                    img_y = mm_to_points(img_y_mm)
                    img_larg = mm_to_points(img_larg_mm)
                    img_alt = mm_to_points(img_alt_mm)
                    imagem = ImageReader(caminho_imagem)
                    can.drawImage(imagem, img_x, img_y, width=img_larg, height=img_alt)
                    
                    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M")
                    can.setFont(fonte, tamanho_fonte)
                    can.drawCentredString(img_x + img_larg / 2, img_y - (tamanho_fonte + 5), data_hora)
                else:
                    if texto_assinatura.strip():
                        can.setFont(fonte, tamanho_fonte)
                        linhas = texto_assinatura.strip().split('\n')
                        base_y = mm_to_points(pos_y_mm)
                        x_texto = mm_to_points(pos_x_mm)

                        for j, linha in enumerate(linhas):
                            y = base_y - (tamanho_fonte + 2) * j
                            can.drawString(x_texto, y, linha)

                        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M")
                        y_data = base_y - (tamanho_fonte + 2) * (len(linhas) + 1)
                        can.drawString(x_texto, y_data, data_hora)

            can.save()
            packet.seek(0)
            overlay = PdfReader(packet)
            pagina.merge_page(overlay.pages[0])
            escritor_pdf.add_page(pagina)

        saida = caminho_pdf_entrada.replace(".pdf", "_assinado.pdf")
        with open(saida, 'wb') as f:
            escritor_pdf.write(f)

        try: os.remove(caminho_pdf_entrada)
        except: pass
        
        return saida
    except Exception as e:
        raise Exception(f"Erro na assinatura: {e}")

def abrir_dialogo_assinatura():
    janela = Toplevel(app)
    janela.title("Assinador Digital de PDFs")
    janela.geometry("600x700")
    janela.configure(bg="white")
    janela.transient(app)
    janela.grab_set()

    # Cabeçalho
    header = tk.Frame(janela, bg=COR_PRIMARIA, height=60)
    header.pack(fill='x')
    header.pack_propagate(False)
    tk.Label(header, text="✒️ ASSINAR DOCUMENTO PDF", font=("Segoe UI", 14, "bold"), bg=COR_PRIMARIA, fg="white").pack(pady=15)

    path_var = StringVar()
    img_path_var = StringVar()
    
    # Estilo específico para LabelFrames
    style_lf = ttk.Style()
    style_lf.configure("Custom.TLabelframe", background="white")
    style_lf.configure("Custom.TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground=COR_SECUNDARIA, background="white")

    main_scroll = tk.Frame(janela, bg="white", padx=20, pady=10)
    main_scroll.pack(fill="both", expand=True)

    # 1. Seleção do Arquivo
    f_frame = ttk.LabelFrame(main_scroll, text=" 1. Seleção do Arquivo ", style="Custom.TLabelframe")
    f_frame.pack(fill='x', pady=10)
    inner_f = tk.Frame(f_frame, bg="white", padx=10, pady=10)
    inner_f.pack(fill='x')
    
    ttk.Entry(inner_f, textvariable=path_var).pack(side='left', fill='x', expand=True, ipady=3)
    tk.Button(inner_f, text="🔍 Buscar PDF", font=("Segoe UI", 9), bg=COR_DESTAQUE, fg="white", relief="flat",
              command=lambda: path_var.set(filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")]))).pack(side='right', padx=5)

    # 2. Texto da Assinatura
    txt_frame = ttk.LabelFrame(main_scroll, text=" 2. Texto da Assinatura ", style="Custom.TLabelframe")
    txt_frame.pack(fill='x', pady=10)
    inner_txt = tk.Frame(txt_frame, bg="white", padx=10, pady=10)
    inner_txt.pack(fill='x')
    
    txt_area = tk.Text(inner_txt, height=3, font=("Segoe UI", 10), bg="#f8f9fa", relief="flat", highlightthickness=1, highlightbackground="#dcdde1")
    txt_area.insert("1.0", "Wesley Raimundo\nAnalista Fiscal Sênior")
    txt_area.pack(fill='x')

    # 3. Posição e Configurações
    pos_frame = ttk.LabelFrame(main_scroll, text=" 3. Configurações de Posição (mm) ", style="Custom.TLabelframe")
    pos_frame.pack(fill='x', pady=10)
    inner_pos = tk.Frame(pos_frame, bg="white", padx=10, pady=10)
    inner_pos.pack(fill='x')
    
    tk.Label(inner_pos, text="X (Horizontal):", bg="white").pack(side='left', padx=5)
    en_x = ttk.Entry(inner_pos, width=8); en_x.insert(0, "170"); en_x.pack(side='left', padx=5)
    
    tk.Label(inner_pos, text="Y (Vertical):", bg="white").pack(side='left', padx=5)
    en_y = ttk.Entry(inner_pos, width=8); en_y.insert(0, "120"); en_y.pack(side='left', padx=5)
    
    tk.Label(inner_pos, text="Tam. Fonte:", bg="white").pack(side='left', padx=5)
    en_size = ttk.Entry(inner_pos, width=5); en_size.insert(0, "8"); en_size.pack(side='left', padx=5)

    # 4. Imagem Opcional
    img_frame = ttk.LabelFrame(main_scroll, text=" 4. Imagem da Assinatura (Opcional) ", style="Custom.TLabelframe")
    img_frame.pack(fill='x', pady=10)
    inner_img = tk.Frame(img_frame, bg="white", padx=10, pady=10)
    inner_img.pack(fill='x')
    
    ttk.Entry(inner_img, textvariable=img_path_var).pack(side='left', fill='x', expand=True, ipady=3)
    tk.Button(inner_img, text="🖼️ Buscar Imagem", font=("Segoe UI", 9), bg="#95a5a6", fg="white", relief="flat",
              command=lambda: img_path_var.set(filedialog.askopenfilename())).pack(side='right', padx=5)

    def processar():
        try:
            pdf = path_var.get()
            if not pdf: 
                messagebox.showwarning("Atenção", "Selecione um arquivo PDF.")
                return
            texto = txt_area.get("1.0", "end-1c")
            x = float(en_x.get())
            y = float(en_y.get())
            size = int(en_size.get())
            img = img_path_var.get() or None
            
            res = assinar_pdf_logic(pdf, texto, x, y, "Helvetica", size, "Todas", 
                                  caminho_imagem=img, img_x_mm=x-20, img_y_mm=y, img_larg_mm=50, img_alt_mm=20)
            
            messagebox.showinfo("Sucesso", f"O documento foi assinado com sucesso!\nSalvo em: {res}")
            os.startfile(res)
            janela.destroy()
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    btn_sign = tk.Button(janela, text="🖋️ ASSINAR DOCUMENTO AGORA", font=("Segoe UI", 11, "bold"), 
                        bg=COR_SUCESSO, fg="white", padx=30, pady=15, relief="flat", command=processar, cursor="hand2")
    btn_sign.pack(pady=20)


# ========== CONFIGURAÇÃO DA INTERFACE PRINCIPAL ==========

def configurar_interface():
    """
    Configura e inicializa a interface gráfica principal do sistema.
    Cria a janela principal com design moderno e profissional.
    """
    global app, entry_busca, tree, notebook, gerar_gnre_tab, lbl_status_monitor, lbl_timer_monitor, entrada_nfe, tree_gnre
    
    # ===== CRIAÇÃO DA JANELA PRINCIPAL =====
    app = Tk()
    app.title(f"{TITULO_SISTEMA} - Versão {VERSAO_SISTEMA}")
    
    # Configurar ícone (se disponível)
    # try:
    #     app.iconbitmap(r"caminho\para\icone.ico")
    # except:
    #     pass
    
    # ===== CENTRALIZAÇÃO E DIMENSIONAMENTO =====
    largura_janela = 1200
    altura_janela = 700
    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()
    x = (screen_width / 2) - (largura_janela / 2)
    y = (screen_height / 2) - (altura_janela / 2)
    app.geometry(f"{largura_janela}x{altura_janela}+{int(x)}+{int(y)}")
    app.resizable(True, True)
    
    # ===== CONFIGURAÇÃO DE CORES CORPORATIVAS =====
    app.configure(bg=COR_FUNDO)
    
    # ===== CONFIGURAÇÃO DE ESTILOS MODERNOS =====
    style = ttk.Style()
    style.theme_use("clam")
    
    # Estilo para botões
    style.configure("TButton",
                   font=("Segoe UI", 10),
                   padding=8,
                   background=COR_PRIMARIA,
                   foreground="white",
                   borderwidth=0)
    style.map("TButton",
             background=[("active", COR_DESTAQUE)])
    
    # Estilo para labels
    style.configure("TLabel",
                   font=("Segoe UI", 10),
                   background=COR_FUNDO,
                   foreground=COR_SECUNDARIA)
    
    # Estilo para notebook (abas)
    style.configure("TNotebook",
                   background=COR_FUNDO,
                   borderwidth=0)
    style.configure("TNotebook.Tab",
                   font=("Segoe UI", 10, "bold"),
                   padding=[20, 10],
                   background=COR_SECUNDARIA,
                   foreground="white")
    style.map("TNotebook.Tab",
             background=[("selected", COR_DESTAQUE)],
             foreground=[("selected", "white")])

    # Estilo para botões de sucesso (Verde)
    style.configure("Success.TButton",
                   font=("Segoe UI", 9, "bold"),
                   foreground="white",
                   background=COR_SUCESSO)
    style.map("Success.TButton",
             background=[("active", COR_DESTAQUE)])
    
    # Estilo para Treeview (tabelas)
    style.configure("Treeview",
                   font=("Segoe UI", 9),
                   rowheight=25,
                   background="white",
                   fieldbackground="white",
                   foreground=COR_SECUNDARIA)
    style.configure("Treeview.Heading",
                   font=("Segoe UI", 10, "bold"),
                   background=COR_PRIMARIA,
                   foreground="white",
                   relief="flat")
    style.map("Treeview.Heading",
             background=[("active", COR_DESTAQUE)])
    
    # ===== CABEÇALHO DO SISTEMA =====
    frame_header = tk.Frame(app, bg=COR_PRIMARIA, height=80)
    frame_header.pack(fill=tk.X, side=tk.TOP)
    frame_header.pack_propagate(False)
    
    # Título principal
    label_titulo = tk.Label(
        frame_header,
        text="📋 SISTEMA GERADOR DE GNRE",
        font=("Segoe UI", 18, "bold"),
        bg=COR_PRIMARIA,
        fg="white"
    )
    label_titulo.pack(side=tk.LEFT, padx=30, pady=20)
    
    # Informações do sistema
    label_versao = tk.Label(
        frame_header,
        text=f"Versão {VERSAO_SISTEMA} | ",
        font=("Segoe UI", 9),
        bg=COR_PRIMARIA,
        fg="#ecf0f1"
    )
    label_versao.pack(side=tk.LEFT, padx=10)
    
    # Indicador de Monitoramento
    frame_monitor = tk.Frame(frame_header, bg=COR_PRIMARIA)
    frame_monitor.pack(side=tk.RIGHT, padx=30)
    
    conf_res = get_configuracoes()
    app.monitoramento_ativo = tk.BooleanVar(value=bool(conf_res.get("AUTO_IMPORT_XML", 1)))
    app.organizador_ativo = tk.BooleanVar(value=bool(conf_res.get("AUTO_ORGANIZAR_XML", 1)))
    app.tempo_organizador = 60
    app.tempo_restante_monitor = 30
    
    chk_monitor = tk.Checkbutton(
        frame_monitor,
        text="Auto-Importar XML (Pasta XML)",
        variable=app.monitoramento_ativo,
        font=("Segoe UI", 9, "bold"),
        bg=COR_PRIMARIA,
        fg="white",
        selectcolor=COR_SECUNDARIA,
        activebackground=COR_PRIMARIA,
        activeforeground=COR_DESTAQUE,
        cursor="hand2"
    )
    chk_monitor.pack(side=tk.RIGHT)
    
    lbl_status_monitor = tk.Label(
        frame_monitor,
        text="●",
        font=("Segoe UI", 12),
        bg=COR_PRIMARIA,
        fg="gray"
    )
    lbl_status_monitor.pack(side=tk.RIGHT, padx=5)
    
    lbl_timer_monitor = tk.Label(
        frame_monitor,
        text="(30s)",
        font=("Segoe UI", 8, "bold"),
        bg=COR_PRIMARIA,
        fg="#bdc3c7"
    )
    lbl_timer_monitor.pack(side=tk.RIGHT, padx=(0, 5))
    
    def atualizar_cor_monitor(*args):
        if app.monitoramento_ativo.get():
            lbl_status_monitor.config(fg=COR_SUCESSO)
            # Gatilho imediato ao ativar
            app.tempo_restante_monitor = 0
        else:
            lbl_status_monitor.config(fg="gray")
            
    app.monitoramento_ativo.trace_add("write", atualizar_cor_monitor)
    
    # ===== FUNÇÃO SOBRE =====
    def sobre(event=None):
        """Exibe informações sobre o sistema"""
        sobre_janela = Toplevel(app)
        sobre_janela.title("Sobre o Sistema")
        sobre_janela.geometry("450x300")
        sobre_janela.resizable(False, False)
        sobre_janela.configure(bg="white")
        sobre_janela.transient(app)
        sobre_janela.grab_set()
        
        # Centralizar janela
        sobre_janela.update_idletasks()
        x = (sobre_janela.winfo_screenwidth() // 2) - (450 // 2)
        y = (sobre_janela.winfo_screenheight() // 2) - (300 // 2)
        sobre_janela.geometry(f"+{x}+{y}")
        
        # Cabeçalho
        frame_sobre_header = tk.Frame(sobre_janela, bg=COR_PRIMARIA, height=60)
        frame_sobre_header.pack(fill=tk.X)
        frame_sobre_header.pack_propagate(False)
        
        tk.Label(
            frame_sobre_header,
            text="ℹ️ Informações do Sistema",
            font=("Segoe UI", 14, "bold"),
            bg=COR_PRIMARIA,
            fg="white"
        ).pack(pady=15)
        
        # Conteúdo
        frame_sobre_content = tk.Frame(sobre_janela, bg="white", padx=30, pady=20)
        frame_sobre_content.pack(fill=tk.BOTH, expand=True)
        
        info_text = f"""
Sistema Gerador de GNRE
Versão: {VERSAO_SISTEMA}
Última Atualização: 05/02/2026

Desenvolvido por: Wesley Raimundo
Empresa: Dinatécnica

Descrição:
Sistema completo para geração, gerenciamento
e controle de GNREs (Guia Nacional de
Recolhimento de Tributos Estaduais).
        """
        
        tk.Label(
            frame_sobre_content,
            text=info_text,
            font=("Segoe UI", 10),
            bg="white",
            fg=COR_SECUNDARIA,
            justify=tk.LEFT
        ).pack(pady=10)
        
        # Botão fechar
        ttk.Button(
            frame_sobre_content,
            text="Fechar",
            command=sobre_janela.destroy
        ).pack(pady=10)
    
    # ===== RODAPÉ =====
    frame_footer = tk.Frame(app, bg=COR_SECUNDARIA, height=40)
    frame_footer.pack(fill=tk.X, side=tk.BOTTOM)
    frame_footer.pack_propagate(False)
    
    footer_label = tk.Label(
        frame_footer,
        text="© 2026  - Desenvolvido por Wesley Raimundo",
        font=("Segoe UI", 9),
        bg=COR_SECUNDARIA,
        fg="white",
        cursor="hand2"
    )
    footer_label.pack(side=tk.RIGHT, padx=20, pady=10)
    footer_label.bind("<Button-1>", sobre)
    
    # ===== NOTEBOOK (ABAS) =====
    notebook = ttk.Notebook(app)
    notebook.pack(fill='both', expand=True, padx=10, pady=10)

    # 1. ABA DASHBOARD (NOVA)
    dashboard_tab = criar_aba_dashboard(notebook)
    notebook.add(dashboard_tab, text="📊 Dashboard")

    # 2. ABA CONSULTA
    main_tab = ttk.Frame(notebook)
    notebook.add(main_tab, text="🔍 Consulta")
    

    import os
    import re
    import sqlite3
    from tkinter import messagebox

    def associar_xmls_por_pasta_raiz():
        DB_PATH = "DADOS_GNRE.db"
        PASTA_RAIZ = "S:/NFE/"  # ou use filedialog.askdirectory()

        if not os.path.exists(PASTA_RAIZ):
            messagebox.showerror("Erro", f"Pasta não encontrada: {PASTA_RAIZ}")
            return

        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()

            padrao = re.compile(r"S3nf0(\d+)", re.IGNORECASE)
            atualizados = 0

            for raiz, _, arquivos in os.walk(PASTA_RAIZ):
                for arquivo in arquivos:
                    match = padrao.match(arquivo)
                    if match and arquivo.lower().endswith(".xml"):
                        numero_nf = match.group(1)
                        caminho_completo = os.path.join(raiz, arquivo)

                        cursor.execute('SELECT "Nº_NFE" FROM DADOS_GNRE WHERE "Nº_NFE" = ?', (numero_nf,))
                        resultado = cursor.fetchone()

                        if resultado:
                            cursor.execute('UPDATE DADOS_GNRE SET CAMINHO_XML = ? WHERE "Nº_NFE" = ?', (caminho_completo, numero_nf))
                            print(f"Atualizado: {numero_nf} -> {caminho_completo}")
                            atualizados += 1

            conn.commit()
            messagebox.showinfo("Concluído", f"XMLs associados com sucesso!\nTotal atualizados: {atualizados}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        finally:
            conn.close()

    # Set up real-time database updates
      

    # Initialize the main tab with real-time updates
    notebook.add(main_tab, text="Consulta")
    
    entry_busca = ttk.Entry(main_tab, width=50)
    entry_busca.pack(pady=10)
    ttk.Button(main_tab, text="Buscar PDF", command=pesquisar_pdf).pack(pady=5)
    columns = ("nfe", "emissao", "uf", "razao", "total", "fecp", "st", "status", "protocolo")
    tree = ttk.Treeview(main_tab, columns=columns, show="headings", selectmode="browse")
    
    # Definição dos Cabeçalhos e Larguras
    tree.heading("nfe", text="Nº NF-e")
    tree.heading("emissao", text="Emissão")
    tree.heading("uf", text="UF")
    tree.heading("razao", text="Razão Social / Cliente")
    tree.heading("total", text="Valor Total")
    tree.heading("fecp", text="FCP")
    tree.heading("st", text="ICMS ST")
    tree.heading("status", text="Status")
    tree.heading("protocolo", text="Protocolo")
    
    tree.column("nfe", width=80, anchor="center")
    tree.column("emissao", width=90, anchor="center")
    tree.column("uf", width=40, anchor="center")
    tree.column("razao", width=300, anchor="w")
    tree.column("total", width=100, anchor="e")
    tree.column("fecp", width=90, anchor="e")
    tree.column("st", width=90, anchor="e")
    tree.column("status", width=120, anchor="center")
    tree.column("protocolo", width=120, anchor="center")
    
    # Barra de Rolagem
    scrollbar = ttk.Scrollbar(main_tab, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    tree.pack(pady=10, fill="both", expand=True)

    # --- MENU DE CONTEXTO (CLIQUE DIREITO) ---
    menu_contexto = Menu(app, tearoff=0)
    menu_contexto.add_command(label="📧 Enviar E-mail Selecionado", 
                             command=lambda: enviar_email_com_anexo(tree, tree.selection()[0]))
    
    def popup_menu(event):
        item = tree.identify_row(event.y)
        if item:
            tree.selection_set(item)
            menu_contexto.post(event.x_root, event.y_root)

    tree.bind("<Button-3>", popup_menu)
            # Configura tags para linhas da tabela
# Dentro de configurar_interface, após a criação de 'tree':
# ...
    tree.tag_configure("anexado", background="lightgreen")
    tree.tag_configure("nao_anexado", background="lightcoral")
    tree.tag_configure("cancelada", background="lightgrey") # Adicionar esta linha!
    atualizar_tabela2()
# ...
    
    
    # ========== CRIAÇÃO DA BARRA DE MENUS ==========
    menubar = Menu(app)

    # ========== MENU ARQUIVO ==========
    menu_arquivo = Menu(menubar, tearoff=0)
    
    menu_arquivo.add_command(label="Importar Arquivos XML de NF-e", 
                            command=importar_xmls)
    menu_arquivo.add_command(label="Importar E-mails de Arquivo", 
                            command=importar_emails_por_arquivo)
    
    menu_arquivo.add_separator()
    
    menu_arquivo.add_command(label="📁 Organizar Comprovantes", 
                            command=organizar_comprovantes)
    
    menu_arquivo.add_separator()
    
    menu_arquivo.add_command(label="Criar Backup do Banco de Dados", 
                            command=criar_backup)
    menu_arquivo.add_command(label="Backup na Nuvem (Google Drive)", 
                            command=backup_google_drive)
    
    menubar.add_cascade(label="Arquivo", menu=menu_arquivo)

    # ========== MENU ASSOCIAÇÃO ==========
    menu_associacao = Menu(menubar, tearoff=0)
    
    # Submenu: Associar PDFs
    submenu_pdfs = Menu(menu_associacao, tearoff=0)
    submenu_pdfs.add_command(label="Associar PDFs por Número de NF-e", 
                            command=associar_pdfs_nf)
    submenu_pdfs.add_command(label="Selecionar Pasta de PDFs", 
                            command=selecionar_pasta_pdf)
    
    menu_associacao.add_cascade(label="Associar PDFs", menu=submenu_pdfs)
    
    # Submenu: Associar XMLs
    submenu_xmls = Menu(menu_associacao, tearoff=0)
    submenu_xmls.add_command(label="Associar XMLs por Pasta Raiz", 
                            command=associar_xmls_por_pasta_raiz)
    submenu_xmls.add_command(label="Selecionar Pasta de XMLs", 
                            command=selecionar_pasta_xml)
    
    menu_associacao.add_cascade(label="Associar XMLs", menu=submenu_xmls)
    
    menu_associacao.add_separator()
    
    # Associar CC-e
    menu_associacao.add_command(label="Selecionar Pasta de CC-e", 
                                command=selecionar_pasta_cce)
    
    menubar.add_cascade(label="Associação", menu=menu_associacao)

    # ========== MENU E-MAILS ==========
    menu_emails = Menu(menubar, tearoff=0)
    
    menu_emails.add_command(label="Cadastrar E-mail de Cliente Manualmente", 
                           command=cadastrar_email_cliente)
    menu_emails.add_command(label="Atualizar E-mails por Código de Parte", 
                           command=atualizar_emails_por_cod_part)
    
    menubar.add_cascade(label="E-mails", menu=menu_emails)

    # ========== MENU BANCO DE DADOS ==========
    menu_bd = Menu(menubar, tearoff=0)
    
    menu_bd.add_command(label="Atualizar Banco de Dados", 
                       command=atualizar_todas_as_tabelas_e_abas)
    
    menubar.add_cascade(label="Banco de Dados", menu=menu_bd)

    # ========== MENU CONFIGURAÇÕES ==========
    menu_configuracoes = Menu(menubar, tearoff=0)
    
    menu_configuracoes.add_command(label="Configurar Pastas do Sistema", 
                                  command=abrir_dialogo_configuracoes_pastas)
    
    menubar.add_cascade(label="Configurações", menu=menu_configuracoes)

    app.config(menu=menubar)
    
    # ===== CONFIGURAÇÃO DAS COLUNAS DA TABELA =====
    # Ajustar alinhamento e largura das colunas
    for col in columns:
        tree.heading(col, text=col, anchor='center')
        # Alinhamento centralizado para melhor visualização
        tree.column(col, width=150, anchor='center')
    tree.pack(pady=5, fill='both', expand=True)
    
            # Criar aba - Gerar GNRE
    gerar_gnre_tab = ttk.Frame(notebook)
    notebook.add(gerar_gnre_tab, text="Gerar GNRE")


configurar_interface()
executar_monitoramento_xml()
executar_organizador_xml()
def formatar_valor(valor):
    """Converte o valor para o formato brasileiro."""
    return f"{valor:,.2f}".replace(",", "TEMP").replace(".", ",").replace("TEMP", ".")
def formatar_valor(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
def data_por_extenso(data):
    meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    return f"{data.day} de {meses[data.month - 1]} de {data.year}"
def formatar_valor(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
def data_por_extenso(data):
    meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    return f"{data.day} de {meses[data.month - 1]} de {data.year}"


def copiar_obs_gnre():
    selecionado = tree_gnre.selection()
    if not selecionado:
        messagebox.showwarning("Atenção", "Selecione uma nota na lista para copiar a Observação.")
        return

    # O campo OBS GNRE é a quinta coluna na nova tree_gnre (índice 4)
    obs_gnre = tree_gnre.item(selecionado[0], 'values')[4]

    pyperclip.copy(obs_gnre)
    messagebox.showinfo("Sucesso", "Observação copiada para a área de transferência!")

def listar_nfe_sem_caminho_pdf():
    tree_gnre.delete(*tree_gnre.get_children())  # limpa a tabela antes

    entrada = entrada_nfe.get().strip()

    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()

    if entrada:  # Se o usuário digitou alguma coisa
        numeros = [num.strip() for num in entrada.split(",") if num.strip()]
        placeholders = ','.join('?' for _ in numeros)
        # FILTRA: Não canceladas, com COD_PART preenchido e não marcadas como SEM ST
        query = f"SELECT * FROM DADOS_GNRE WHERE `Nº_NFE` IN ({placeholders}) AND CANCELADA = 0 AND (COD_PART IS NOT NULL AND COD_PART != '') AND (CAMINHO_PDF IS NULL OR CAMINHO_PDF != 'SEM ST') ORDER BY DT_EMISSÃO DESC, Nº_NFE DESC" 
        cursor.execute(query, numeros)
    else:
        # 3. NÃO estão marcadas como SEM ST na coluna CAMINHO_PDF
        # 4. Têm COD_PART preenchido
        # Ordena por data decrescente
        cursor.execute("SELECT * FROM DADOS_GNRE WHERE (CAMINHO_PDF IS NULL OR CAMINHO_PDF = '') AND CANCELADA = 0 AND (COD_PART IS NOT NULL AND COD_PART != '') ORDER BY DT_EMISSÃO DESC, Nº_NFE DESC")

    resultados = cursor.fetchall()
    conn.close()

    for row_data in resultados:
        # Mapeia os dados da tupla para a ordem da nova tree_gnre:
        # ["Nº NFE", "UF", "VALOR TOTAL", "CLIENTE", "OBS GNRE"]
        # Índices: [0, 4, 33, 3, 18]
        nova_linha = [
            row_data[0], # Nº NFE
            row_data[4], # UF
            row_data[33], # VALOR_TOTAL_GNRE
            row_data[3], # RAZÃO_SOCIAL_TOMADOR
            row_data[18] # OBS_GNRE
        ]
        tree_gnre.insert("", "end", values=nova_linha, iid=row_data[0]) # Usa o Nº NFE como iid para referência futura


# Função para excluir item da tabela
def excluir_item_da_tabela():
    """Excluir item da tabela"""
    selected_item = tree_gnre.selection()[0]
    tree_gnre.delete(selected_item)

ano_atual = datetime.now().year
mes_atual = datetime.now().strftime("%m - %B").lower()  # Formato de mês: 01 - Janeiro, 02 - Fevereiro, 03 - Março, 04 - Abril, 05 - Maio, 06 - Junho, 07 - Julho, 08 - Agosto, 09 - Setembro, 10 - Outubro, 11 - Novembro, 12 - Dezembro
def salvar_pdf():
    """
    Salva o grid em um arquivo PDF (CI), usando o mapeamento de índices para
    garantir que os dados da tabela estejam na ordem e formato corretos.
    """
    
    # --- AJUSTE ESTES ÍNDICES CONFORME A ORDEM ATUAL DA SUA TABELA DADOS_GNRE ---
    # Se você adicionou 'CANCELADA' no final, os índices altos (como 33, 40) podem ter mudado.
    INDEX_MAP = {
        'Nº_NFE': 0,           # Ex: 0
        'DT_EMISSÃO': 1,       # Ex: 1
        'RAZÃO_SOCIAL_TOMADOR': 3, # Ex: 3 (CLIENTE)
        'UF_TOMADOR': 4,       # Ex: 4 (UF)
        'PROTOCOLO_ICMS': 12,  # Ex: 12 (PROTOCOLO)
        'VALOR_TOTAL_GNRE': 33,# Ex: 33 (VALOR)
    }
    # -------------------------------------------------------------------------
    
    nfe_numeros = []
    for row_id in tree_gnre.get_children():
        nfe_numero = tree_gnre.item(row_id, 'values')[0]
        nfe_numeros.append(nfe_numero)

    if not nfe_numeros:
        messagebox.showwarning("Aviso", "Não há dados na tabela para exportar.")
        return

    try:
        # Configuração de Locale (mantido)
        try:
            locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
        except locale.Error:
            try:
                locale.setlocale(locale.LC_TIME, 'pt_BR')
            except locale.Error:
                messagebox.showwarning("Aviso", "Não foi possível configurar o locale para português.")
        
        # --- DEFINIÇÃO DO CAMINHO ---
        gnre_root = r"D:\01 - SISTEMAS E TRIBUTARIO\GUIA ST" 
        now = datetime.now()
        ano = now.strftime("%Y")
        mes_numero = now.strftime("%m")
        mes_nome = now.strftime("%B").capitalize()
        pasta_destino = os.path.join(gnre_root, ano, f"{mes_numero} - {mes_nome}", "CI")

        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)
            
        data_atual = now.strftime("%d-%m-%Y")
        nome_arquivo = f"CI GNRE {data_atual}.pdf"
        filepath = os.path.join(pasta_destino, nome_arquivo)
        # --- FIM DA DEFINIÇÃO DO CAMINHO ---

        # 1. CONEXÃO E BUSCA DOS DADOS COMPLETOS
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()
        
        placeholders = ','.join('?' for _ in nfe_numeros)
        cursor.execute(f"SELECT * FROM DADOS_GNRE WHERE `Nº_NFE` IN ({placeholders})", nfe_numeros)
        notas_completas = cursor.fetchall()
        conn.close()
        
        if not notas_completas:
            messagebox.showwarning("Aviso", "Nenhum dado completo encontrado no banco para exportar.")
            return

        # 2. INÍCIO DA GERAÇÃO DO PDF
        doc = SimpleDocTemplate(filepath, pagesize=A4)
        elements = []
        styles = getSampleStyleSheet()
        
        # ... (BLOCO DE CABEÇALHO E IMAGEM) ...
        header = "COMUNICAÇÃO INTERNA"
        date_today = datetime.today().strftime('%d/%m/%Y')
        header_style = ParagraphStyle(name='HeaderStyle', fontSize=10, alignment=2)
        header_paragraph = Paragraph(header, header_style)
        
        url_imagem = "https://www.dinatecnica.com.br/assets/uploads/media-uploader/logo-dinatecnica1713869009.png"
        caminho_local = r"C:\Athena\62522453000135.bmp"

        img = None 
        try:
            response = requests.get(url_imagem, timeout=5)
            response.raise_for_status() 
            image_data = BytesIO(response.content)
            img = Image(image_data)
        except Exception:
            try:
                img = Image(caminho_local)
            except Exception:
                pass

        if img:
            img_width, img_height = 5 * cm, 1 * cm
            img.drawWidth = img_width
            img.drawHeight = img_height
            header_table_data = [[img, header_paragraph]]
        else:
            header_table_data = [["", header_paragraph]]
            
        header_table = Table(header_table_data, colWidths=[4 * cm, 10 * cm])
        header_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
        elements.append(header_table)
        
        # INFORMAÇÕES DO REMETENTE (mantido)
        formatted_date = now.strftime("%d%m%Y")
        formatted_time = now.strftime("%H%M%S")
        wes_number = f"Nº WES{formatted_date}{formatted_time}"
        info_style = ParagraphStyle(name='InfoStyle', fontSize=8, leading=14, fontName='Helvetica-Bold')
        user_name = getpass.getuser()
        elements.append(Paragraph(wes_number, info_style))
        elements.append(Paragraph(f"De: {user_name.upper()}", info_style))
        elements.append(Paragraph("Para: Marcilio GA/Enéas Financeiro", info_style))
        elements.append(Paragraph("C/C: Edson / Contabilidade / Fiscal", info_style))
        elements.append(Paragraph("ASSUNTO: RECOLHIMENTO DE SUBSTITUIÇÃO TRIBUTÁRIA POR OPERAÇÃO", info_style))
        elements.append(Spacer(1, 12))
        request_paragraph_style = ParagraphStyle(name='RequestStyle', fontSize=10, leading=16, textColor=colors.black)
        request_paragraph = Paragraph("Favor providenciar o pagamento nesta data, conforme abaixo:", request_paragraph_style)
        elements.append(request_paragraph)
        elements.append(Spacer(0, 0))
        
        # 3. CRIAÇÃO DA TABELA DE DADOS
        
        colunas_gnre = ["PROTOCOLO", "Nº NFE", "UF", "VALOR", "EMISSÃO", "VENCTO", "CLIENTE"]
        data = []
        header_row = ["#", *colunas_gnre]
        data.append(header_row)
        total_icms_st = 0
        
        for row in notas_completas: 
            
            # Use os índices do INDEX_MAP
            idx_protocolo = INDEX_MAP['PROTOCOLO_ICMS']
            idx_nfe = INDEX_MAP['Nº_NFE']
            idx_uf = INDEX_MAP['UF_TOMADOR']
            idx_valor = INDEX_MAP['VALOR_TOTAL_GNRE']
            idx_emissao = INDEX_MAP['DT_EMISSÃO']
            idx_cliente = INDEX_MAP['RAZÃO_SOCIAL_TOMADOR'] 
            
            # 1. #
            nova_linha = [len(data)]
            
            # 2. PROTOCOLO
            nova_linha.append(str(row[idx_protocolo]))
            
            # 3. Nº NFE
            nova_linha.append(str(row[idx_nfe]))
            
            # 4. UF
            nova_linha.append(str(row[idx_uf]))
            
            # 5. VALOR (com formatação e soma)
            valor_float = 0.0
            try:
                valor_original = row[idx_valor]
                valor_float = float(valor_original)
                valor_formatado = formatar_valor(valor_float)
                total_icms_st += valor_float
            except (ValueError, TypeError):
                valor_formatado = "0,00"
                
            nova_linha.append(valor_formatado)
            
            # 6. EMISSÃO
            data_emissao_str = "Data inválida"
            try:
                data_emissao = datetime.strptime(row[idx_emissao], '%Y-%m-%d').strftime('%d/%m/%Y')
                data_emissao_str = data_emissao
            except (ValueError, TypeError):
                pass
            nova_linha.append(data_emissao_str)
            
            # 7. VENCTO (Data de Hoje)
            nova_linha.append(datetime.today().strftime('%d/%m/%Y'))
            
            # 8. CLIENTE
            nova_linha.append(str(row[idx_cliente]))
                
            # Adiciona a linha (deve ter 8 itens: # + 7 colunas)
            if len(nova_linha) == len(header_row):
                data.append(nova_linha)
            # Se a linha estiver incompleta, o erro de "tuple index out of range" no PDF foi evitado,
            # mas indica que o mapeamento de índices está errado, e esta linha é ignorada.


        if len(data) <= 1: # Só tem o cabeçalho
            messagebox.showwarning("Aviso", "Não há dados válidos para exportar (Verifique o mapeamento de índices).")
            return
            
        # 4. FORMATAÇÃO E CONSTRUÇÃO FINAL
        table = Table(data)
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.beige, colors.whitesmoke])
        ])
        table.setStyle(style)
        elements.append(Spacer(0, 12))
        elements.append(table)
        
        # Total
        valor_total_icms_st = formatar_valor(total_icms_st)
        valor_icms_paragraph_style = ParagraphStyle(name='ValorICMSStyle', fontSize=10, leading=12)
        valor_icms_paragraph = Paragraph(f"Valor do ICMS ST do dia {date_today}: {valor_total_icms_st}", valor_icms_paragraph_style)
        elements.append(Spacer(1, 12))
        elements.append(valor_icms_paragraph)
        
        # Rodapé e Assinaturas (mantido)
        footer_style = ParagraphStyle(name='FooterStyle', fontSize=10, alignment=2)
        data_impresso = datetime.now().strftime('%d.%m.%y %H:%M')
        footer_info = f"Embu das Artes, {data_impresso}"
        footer_paragraph = Paragraph(footer_info, footer_style)
        elements.append(Spacer(1, 12))
        elements.append(footer_paragraph)
        
        # Assinaturas
        user_name = getpass.getuser()
        signature_paragraph = Paragraph("Faturamento:___________________", styles['Normal'])
        elements.append(signature_paragraph)
        user_paragraph = Paragraph(user_name, styles['Normal'])
        elements.append(user_paragraph)
        
        line_height = 12
        space = Spacer(1, 5 * line_height)
        elements.append(space)
        finance_paragraph = Paragraph("Financeiro: __________________", styles['Normal'])
        elements.append(finance_paragraph)

        # Constrói o documento
        doc.build(elements)

        # Copia o caminho para a área de transferência
        pyperclip.copy(filepath)

        messagebox.showinfo("Sucesso", f"O arquivo PDF foi gerado com sucesso!\nO caminho foi copiado para a área de transferência:\n{filepath}")

        if messagebox.askyesno("Abrir PDF", "Deseja abrir o arquivo agora?"):
            os.startfile(filepath)
            
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao salvar o PDF: {e}")
# Função para buscar dados da NFE
def buscar_dados_nfe(event=None):
    
    numero_nfe = entrada_nfe.get()
    if numero_nfe:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM DADOS_GNRE WHERE Nº_NFE = ?", (numero_nfe,))
        resultado = cursor.fetchall()
        conn.close()
        
        # Limpa a tree_gnre antes de inserir novos resultados
        for row in tree_gnre.get_children():
            tree_gnre.delete(row)
            
        if resultado:
            if resultado[0][34]:  # Verifica se o campo caminho pdf está preenchido
                messagebox.showwarning("GNRE já gerado", f"O PDF da Nota Fiscal {numero_nfe} já foi gerado anteriormente.")
            else:
                for row_data in resultado:
                    # Mapeia os dados da tupla para a ordem da nova tree_gnre:
                    # ["Nº NFE", "UF", "VALOR TOTAL", "CLIENTE", "OBS GNRE"]
                    # Índices: [0, 4, 33, 3, 18]
                    nova_linha = [
                        row_data[0], # Nº NFE
                        row_data[4], # UF
                        row_data[33], # VALOR_TOTAL_GNRE
                        row_data[3], # RAZÃO_SOCIAL_TOMADOR
                        row_data[18] # OBS_GNRE
                    ]
                    # Insere os dados encontrados na tabela
                    tree_gnre.insert("", "end", values=nova_linha, iid=row_data[0]) # Usa o Nº NFE como iid para referência futura
        else:
            messagebox.showwarning("Dados não encontrados", f"Não foram encontrados dados para o número da Nota Fiscal: {numero_nfe}")
    else:
        messagebox.showwarning("Erro", "Por favor, insira um número de Nota Fiscal válido.")


    # Focar e selecionar o campo de entrada para novo número
    entrada_nfe.focus()  # Define o foco no campo de entrada
    entrada_nfe.select_range(0, tk.END)  # Seleciona todo o texto no campo
style = ttk.Style()
style.theme_use('clam')
style.configure("Treeview.Heading", font=("Helvetica", 10), foreground="black", background="grey")
style.configure("Treeview", font=("Helvetica", 9), rowheight=25)
style.map("Treeview.Heading", background=[("active", "grey")])

ttk.Label(gerar_gnre_tab, text="Número da Nota Fiscal:").pack(pady=5)

entrada_nfe = ttk.Entry(gerar_gnre_tab)
entrada_nfe.pack(pady=5)
# Bind o evento Enter para chamar a função de busca
entrada_nfe.bind("<Return>", buscar_dados_nfe)  # Chama buscar_dados_nfe quando Enter é pressionado
# Frame para agrupar os botões
button_frame = ttk.Frame(gerar_gnre_tab) 
button_frame.pack(pady=10)
# Botão para salvar o PDF
ttk.Button(button_frame, text="Salvar PDF", command=salvar_pdf).grid(row=0, column=1, padx=5)
ttk.Button(button_frame, text="Listar NF-e", command=listar_nfe_sem_caminho_pdf).grid(row=0, column=2, padx=5)
ttk.Button(button_frame, text="Excluir", command=excluir_item_da_tabela).grid(row=0, column=3, padx=5)
# Adicione o novo botão AQUI:
ttk.Button(button_frame, text="Renomear Lote PDF", command=renomear_lote_pdf_gui).grid(row=0, column=4, padx=5)
ttk.Button(button_frame, text="Assinar PDF", command=abrir_dialogo_assinatura).grid(row=0, column=5, padx=5)
ttk.Button(button_frame, text="Marcar Cancelada", command=marcar_como_cancelada).grid(row=0, column=6, padx=5) # Novo botão
ttk.Button(button_frame, text="Copiar OBS", command=copiar_obs_gnre).grid(row=0, column=7, padx=5)
columns_gnre = ["Nº NFE", "UF", "VALOR TOTAL", "CLIENTE", "OBS GNRE"]
tree_gnre = ttk.Treeview(gerar_gnre_tab, columns=columns_gnre, show="headings")
for col in columns_gnre:
    tree_gnre.heading(col, text=col)
    # Ajuste as larguras para melhor visualização (opcional)
    if col == "OBS GNRE":
        tree_gnre.column(col, width=300, anchor='w')
    elif col == "CLIENTE":
        tree_gnre.column(col, width=200, anchor='w')
    else:
        tree_gnre.column(col, width=100, anchor='center')
tree_gnre.pack(pady=10, fill="both", expand=True)

# Função para gerar o arquivo XML para cada nota selecionada
import webbrowser # Adicione esta linha no início do seu arquivo, junto com os outros imports

# ... (seu código existente)

def gerar_arquivos_gnre_agrupado():
    """
    Função para gerar o arquivo XML para cada nota selecionada e salvar automaticamente,
    buscando os dados completos no banco de dados para evitar o erro de índice.
    """
    
    # 1. COLETAR OS NÚMEROS DE NFE DO GRID VISÍVEL
    nfe_numeros = []
    for row_id in tree_gnre.get_children():
        # O Nº NFE é a primeira coluna do grid, que é o campo que identifica a nota
        nfe_numero = tree_gnre.item(row_id, 'values')[0] 
        nfe_numeros.append(nfe_numero)

    if not nfe_numeros:
        messagebox.showwarning("Atenção", "Nenhuma nota está no grid para gerar o XML.")
        return

    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()

        # 2. BUSCAR OS DADOS COMPLETOS DE TODAS AS NOTAS SELECIONADAS NO BD
        placeholders = ','.join('?' for _ in nfe_numeros)
        
        # BUSCA OS DADOS COMPLETOS PARA GERAR O XML CORRETAMENTE
        cursor.execute(f"SELECT * FROM DADOS_GNRE WHERE Nº_NFE IN ({placeholders})", nfe_numeros)
        notas_completas = cursor.fetchall()

        # BUSCAR OS DADOS DO EMITENTE (necessário para o bloco <contribuinteEmitente>)
        # Pega a primeira linha do BD para obter os dados do emitente (CNPJ, Razão Social, etc.)
        cursor.execute("SELECT * FROM DADOS_GNRE LIMIT 1")
        emitente = cursor.fetchone()
        
        conn.close()

        # Verifica se encontrou dados do emitente
        if not emitente:
            messagebox.showerror("Erro", "Não foi possível recuperar os dados do emitente do banco de dados.")
            return

        # --- Definir o locale e criar o caminho do diretório ---
        try:
            locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
        except locale.Error:
            try:
                locale.setlocale(locale.LC_TIME, 'pt_BR')
            except locale.Error:
                pass # Ignora se não conseguir configurar o locale

        # (Idealmente, use get_configuracoes() aqui, mas mantendo a estrutura atual para substituição direta)
        gnre_root = r"D:\01 - SISTEMAS E TRIBUTARIO\GUIA ST" # Substituir pela variável de configuração em um refactoring futuro
        
        now = datetime.now()
        ano = now.strftime("%Y")
        mes_numero = now.strftime("%m")
        mes_nome = now.strftime("%B").capitalize()
        
        pasta_destino = os.path.join(gnre_root, ano, f"{mes_numero} - {mes_nome}", "LOTE")

        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)

        current_datetime = now.strftime("Lote GNRE %d-%m-%y")
        filename = f"{current_datetime}.xml"
        filepath = os.path.join(pasta_destino, filename)
        # --- FIM DO NOVO CÓDIGO ---

        root = ET.Element("TLote_GNRE", xmlns="http://www.gnre.pe.gov.br", versao="2.00")
        guias = ET.SubElement(root, "guias")
        
        # Agora o loop itera sobre a tupla COMPLETA (notas_completas)
        for nota in notas_completas: 
            # DADOS DA GNRE:
            dados_gnre = ET.SubElement(guias, "TDadosGNRE", versao="2.00")
            ET.SubElement(dados_gnre, "ufFavorecida").text = nota[4]
            ET.SubElement(dados_gnre, "tipoGnre").text = "0"
            
            # CONTRIBUINTE EMITENTE (usa a tupla emitente, que é a linha completa)
            contribuinte_emitente = ET.SubElement(dados_gnre, "contribuinteEmitente")
            identificacao = ET.SubElement(contribuinte_emitente, "identificacao")
            ET.SubElement(identificacao, "CNPJ").text = emitente[23]
            ET.SubElement(contribuinte_emitente, "razaoSocial").text = emitente[24]
            ET.SubElement(contribuinte_emitente, "endereco").text = emitente[25]
            ET.SubElement(contribuinte_emitente, "municipio").text = emitente[26]
            ET.SubElement(contribuinte_emitente, "uf").text = emitente[27]
            ET.SubElement(contribuinte_emitente, "cep").text = emitente[28]
            ET.SubElement(contribuinte_emitente, "telefone").text = emitente[29]
            
            # ITENS GNRE
            itens_gnre = ET.SubElement(dados_gnre, "itensGNRE")
            item = ET.SubElement(itens_gnre, "item")
            ET.SubElement(item, "receita").text = nota[32]
            
            # DOCUMENTO DE ORIGEM
            if nota[4] == "MT":
                ET.SubElement(item, "detalhamentoReceita").text = "000055"
                ET.SubElement(item, "documentoOrigem", tipo=nota[30]).text = nota[31]
            elif nota[4] == "AL" and nota[5] == "9":
                ET.SubElement(item, "documentoOrigem", tipo=nota[30]).text = nota[31]
            elif nota[4] == "PE" and nota[5] == "9":
                ET.SubElement(item, "documentoOrigem", tipo="24").text = nota[31]
            else:
                ET.SubElement(item, "documentoOrigem", tipo=nota[30]).text = nota[31]
                
            ET.SubElement(item, "produto").text = "18"
            
            # REFERÊNCIA
            referencia = ET.SubElement(item, "referencia")
            ET.SubElement(referencia, "periodo").text = "0"
            ET.SubElement(referencia, "mes").text = nota[16]
            ET.SubElement(referencia, "ano").text = datetime.now().strftime("%Y")
            
            ET.SubElement(item, "dataVencimento").text = datetime.now().strftime("%Y-%m-%d")
            
            # VALOR GNRE (Tipo 11 - Principal)
            if nota[4] == "CE":
                ET.SubElement(item, "valor", tipo="11").text = nota[6]
            elif nota[5] == "9":
                ET.SubElement(item, "valor", tipo="11").text = nota[6]
            elif nota[4] == "AL" or nota[4] == "RJ":
                ET.SubElement(item, "valor", tipo="11").text = nota[8]
            else:
                ET.SubElement(item, "valor", tipo="11").text = nota[33]
                
            # VALOR GNRE (Tipo 12 - FECOP)
            if nota[4] == "MG" or nota[4] == "PR" or nota[4] == "MS" or nota[4] == "BA" or nota[4] == "CE" or nota[4] == "RS" or nota[4] == "PE" or nota[4] == "MT" or nota[4] == "AM":
                ET.SubElement(item, "valor", tipo="12").text = "0.00"
            else:
                # O índice [39] é o vFCPUFDest e [7] é o VL_FECP (o seu cálculo original)
                vl_fecp_total = float(nota[7]) + float(nota[39]) if nota[5] == "9" else float(nota[7])
                ET.SubElement(item, "valor", tipo="12").text = f"{vl_fecp_total:.2f}"
                
            ET.SubElement(item, "convenio").text = nota[12]
            
            # CONTRIBUINTE DESTINATÁRIO
            contribuinte_dest = ET.SubElement(item, "contribuinteDestinatario")
            identificacao_dest = ET.SubElement(contribuinte_dest, "identificacao")
            
            # Lógica específica para RJ Não Contribuinte
            if nota[4] == "RJ" and nota[5] == "9":
                ET.SubElement(identificacao_dest, "CNPJ").text = nota[14]
                ET.SubElement(contribuinte_dest, "razaoSocial").text = nota[3]
                ET.SubElement(contribuinte_dest, "municipio").text = nota[40][-5:]
            if not (nota[4] == "RJ" and nota[5] == "9"):
                ET.SubElement(identificacao_dest, "IE").text = nota[15]
                
            # CAMPOS EXTRAS
            campos_extras = ET.SubElement(item, "camposExtras")
            
            # Campo Extra 1 (Código: nota[21])
            campo_extra_1 = ET.SubElement(campos_extras, "campoExtra")
            ET.SubElement(campo_extra_1, "codigo").text = nota[21]
            
            # Lógica de Valor Específica para Campo Extra 1
            valor_campo_extra_1 = datetime.now().strftime("%Y-%m-%d") # Padrão: data de hoje
            if (nota[4] == "AL" and nota[5] == "1") or \
               (nota[4] == "PR" and nota[5] == "9") or \
               (nota[4] == "MS" and nota[5] == "9") or \
               (nota[4] == "BA" and nota[5] == "9"):
                valor_campo_extra_1 = nota[9] # CHAVE_NFE
            elif (nota[4] == "CE" and nota[5] == "9") or \
                 (nota[4] == "RJ" and nota[5] == "9"):
                valor_campo_extra_1 = datetime.now().strftime("%Y-%m-%d")
                
            ET.SubElement(campo_extra_1, "valor").text = valor_campo_extra_1
            
            # Campo Extra 2 (Código: nota[22])
            campo_extra_2 = ET.SubElement(campos_extras, "campoExtra")
            ET.SubElement(campo_extra_2, "codigo").text = nota[22]
            ET.SubElement(campo_extra_2, "valor").text = nota[18] # OBS_GNRE
            
            # VALOR TOTAL E DATA DE PAGAMENTO
            ET.SubElement(dados_gnre, "valorGNRE").text = nota[33]
            ET.SubElement(dados_gnre, "dataPagamento").text = datetime.now().strftime("%Y-%m-%d")

        # Escreve o arquivo XML
        tree = ET.ElementTree(root)
        tree.write(filepath, encoding="utf-8", xml_declaration=True)

        pyperclip.copy(filepath)

        messagebox.showinfo("Sucesso", f"Arquivos GNRE gerados com sucesso!\nSalvo em: {filepath}")

        # Pergunta se quer abrir o site
        resposta = messagebox.askyesno("Abrir Site GNRE", "Deseja abrir o site da GNRE para envio do lote?")
        if resposta:
            webbrowser.open("https://www.gnre.pe.gov.br:444/gnre/v/lote/gerar")

    except Exception as e:
        # Garante que qualquer erro seja capturado e exibido
        messagebox.showerror("Erro", f"Ocorreu um erro ao gerar o XML: {e}")


def enviar_dua_es_webservice(nfe_numeros, thumbprint, cert_path):
    """
    Realiza a emissão de DUA (Documento Único de Arrecadação) para o estado do Espírito Santo
    através do Webservice oficial da SEFAZ-ES.
    Em caso de falha de envio ou retorno não processado, o XML é reenviado
    automaticamente com um novo <xIde> (arquivo "renomeado") para forçar a
    geração de um novo protocolo. Quando um protocolo válido é recebido ele
    é gravado nas colunas PROTOCOLO_ICMS e PROTOCOLO_OBS do banco, garantindo
    que a geração de PDF reflita o novo número.
    """
    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()
        placeholders = ','.join('?' for _ in nfe_numeros)
        cursor.execute(f"SELECT * FROM DADOS_GNRE WHERE Nº_NFE IN ({placeholders})", nfe_numeros)
        notas_completas = cursor.fetchall()
        
        # Coletar dados do emitente para o CNPJ Emissor
        cursor.execute("SELECT CNPJ_EMITENTE FROM DADOS_GNRE LIMIT 1")
        row_emi = cursor.fetchone()
        cnpj_emitente = str(row_emi[0]).strip() if row_emi else ""
        conn.close()

        if not notas_completas:
            return

        sucessos = []
        erros = []

        # helper para atualizar protocolo no banco (usado tanto no primeiro envio quanto no reenvio)
        def atualizar_protocolo(nfe, protocolo):
            try:
                conn_up = sqlite3.connect("DADOS_GNRE.db")
                cur_up = conn_up.cursor()
                cur_up.execute(
                    "UPDATE DADOS_GNRE SET PROTOCOLO_ICMS = ?, PROTOCOLO_OBS = ? WHERE `Nº_NFE` = ?",
                    (protocolo, protocolo, nfe)
                )
                conn_up.commit()
            except Exception:
                pass
            finally:
                conn_up.close()

        # envia o envelope SOAP e retorna o texto de retorno
        def enviar_soap(soap_env, nfe_num, tag=""):
            url = "https://app.sefaz.es.gov.br/WsDua/DuaService.asmx"
            retorno_local = ""
            if thumbprint:
                temp_name = f"dua_es_{nfe_num}{tag}.xml"
                temp_soap = os.path.join(os.environ['TEMP'], temp_name)
                with open(temp_soap, "w", encoding="utf-8") as f:
                    f.write(soap_env)
                ps_script = f"""
                    $ProgressPreference = 'SilentlyContinue'
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    try {{
                        $body = [System.IO.File]::ReadAllText('{temp_soap}')
                        $cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object {{ $_.Thumbprint -eq '{thumbprint.strip()}' }}
                        if (-not $cert) {{ throw "Certificado nao encontrado" }}
                        $headers = @{{ "Content-Type" = "application/soap+xml; charset=utf-8" }}
                        $resp = Invoke-WebRequest -Uri "{url}" -Method Post -Body $body -Certificate $cert -Headers $headers -UseBasicParsing
                        Write-Output $resp.Content
                    }} catch {{
                        Write-Output "ERRO: $($_.Exception.Message)"
                        if ($_.Exception.Response) {{
                            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                            Write-Output "DETALHE: $($reader.ReadToEnd())"
                        }}
                        exit 1
                    }}
                    """
                res = subprocess.run(["powershell", "-Command", ps_script], capture_output=True, text=True)
                retorno_local = res.stdout
                if os.path.exists(temp_soap):
                    os.remove(temp_soap)
            else:
                from cryptography.hazmat.primitives import serialization
                import tempfile
                senha = simpledialog.askstring("Senha", f"Senha do certificado para nota {nfe_num}:", show='*')
                if not senha:
                    return ""  # usuário cancelou; deixamos retorno vazio para tratar
                with open(cert_path, "rb") as f:
                    pfx_data = f.read()
                p_key, cert, adds = pkcs12.load_key_and_certificates(pfx_data, senha.encode(), default_backend())
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pem") as tmp:
                    tmp.write(p_key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.TraditionalOpenSSL, serialization.NoEncryption()))
                    tmp.write(cert.public_bytes(serialization.Encoding.PEM))
                    pem_path = tmp.name
                try:
                    resp = requests.post(url, data=soap_env.encode('utf-8'), 
                                           headers={"Content-Type": "application/soap+xml; charset=utf-8"}, 
                                           cert=pem_path, timeout=30)
                    retorno_local = resp.text
                finally:
                    if os.path.exists(pem_path):
                        os.remove(pem_path)
            return retorno_local

        for nota in notas_completas:
            nfe_num = nota[0]
            try:
                # Mapeamento de Receita para Area/Serviço no ES
                cod_receita = str(nota[32]).strip()
                cArea = "1"
                cServ = "137"
                
                if cod_receita == "100099":
                    cArea = "1"; cServ = "138"
                elif cod_receita == "100129":
                    cArea = "5"; cServ = "3867"
                
                # Validação de dados mínimos
                cnpj_pes = str(nota[14]).strip()
                v_tot = "{:.2f}".format(float(nota[33].replace(",", ".")))
                c_mun = str(nota[26]).strip() # 5 dígitos conforme manual
                if len(c_mun) > 5: c_mun = c_mun[-5:]
                

                # Montar XML de Dados (emisDua) conforme Manual 1.01b
                xml_dados = f"""<emisDua versao="1.01" xmlns="http://www.sefaz.es.gov.br/duae">
                    <tpAmb>1</tpAmb>
                    <cnpjEmi>{cnpj_emitente}</cnpjEmi>
                    <cnpjOrg>27080571000130</cnpjOrg>
                    <cArea>{cArea}</cArea>
                    <cServ>{cServ}</cServ>
                    <cnpjPes>{cnpj_pes}</cnpjPes>
                    <dRef>{datetime.now().strftime('%Y-%m')}</dRef>
                    <dVen>{datetime.now().strftime('%Y-%m-%d')}</dVen>
                    <dPag>{datetime.now().strftime('%Y-%m-%d')}</dPag>
                    <cMun>{c_mun}</cMun>
                    <xInf>{str(nota[18])[:250]}</xInf>
                    <vRec>{v_tot}</vRec>
                    <qtde>1.0000</qtde>
                    <xIde>{nfe_num}</xIde>
                    <fPix>true</fPix>
                </emisDua>"""





                # SOAP Envelope corrigido conforme MANUAL OFICIAL (SOAP 1.2)
                # Namespace correto: http://www.sefaz.es.gov.br/duae
                # O XML interno deve ser passado como string (CDATA ou Escapado)
                soap_env = f"""<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Header>
    <DuaServiceHeader xmlns="http://www.sefaz.es.gov.br/duae">
      <versao>1.01</versao>
    </DuaServiceHeader>
  </soap12:Header>
  <soap12:Body>
    <duaEmissao xmlns="http://www.sefaz.es.gov.br/duae">
      <duaDadosMsg><![CDATA[{xml_dados}]]></duaDadosMsg>
    </duaEmissao>
  </soap12:Body>
</soap12:Envelope>"""

                url = "https://app.sefaz.es.gov.br/WsDua/DuaService.asmx"
                
                # Salvar debug do ENVIO ES
                try:
                    debug_req_es = os.path.join(os.environ["USERPROFILE"], "Desktop", f"DEBUG_ENVIO_ES_{nfe_num}.xml")
                    with open(debug_req_es, "w", encoding="utf-8") as f:
                        f.write(soap_env)
                except:
                    pass

                # ----- novo fluxo: enviar e tratar retorno (com reenvio automático se necessário) -----
                retorno = enviar_soap(soap_env, nfe_num)
                try:
                    debug_path_es = os.path.join(os.environ["USERPROFILE"], "Desktop", f"DEBUG_RESPOSTA_ES_{nfe_num}.xml")
                    with open(debug_path_es, "w", encoding="utf-8") as f:
                        f.write(retorno)
                except:
                    pass

                # interpretar resposta
                cStat_match = re.search(r"(?:cStat&gt;|cStat>|cStat :)(\d+)", retorno)
                cStat = cStat_match.group(1) if cStat_match else ""
                if cStat == "105":
                    match = re.search(r"(?:nDua&gt;|nDua>|nDua :)\s*(\d+)", retorno)
                    nDua = match.group(1) if match else "???"
                    sucessos.append(f"NFe {nfe_num} -> DUA {nDua}")
                    atualizar_protocolo(nfe_num, nDua)
                else:
                    # preparar reenvio com novo xIde
                    xMotivo_match = re.search(r"(?:xMotivo&gt;|xMotivo>|xMotivo :)\s*([^&<]+)", retorno)
                    motivo = xMotivo_match.group(1).strip() if xMotivo_match else ""
                    if not motivo:
                        fault_match = re.search(r"(?:faultstring&gt;|faultstring>|faultstring :)\s*([^&<]+)", retorno)
                        motivo = fault_match.group(1).strip() if fault_match else "Erro desconhecido na resposta"

                    suffix = datetime.now().strftime("%Y%m%d%H%M%S")
                    novo_xide = f"{nfe_num}-{suffix}"
                    xml_dados_retry = re.sub(r"<xIde>.*?</xIde>", f"<xIde>{novo_xide}</xIde>", xml_dados)
                    soap_env_retry = f"""<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Header>
    <DuaServiceHeader xmlns="http://www.sefaz.es.gov.br/duae">
      <versao>1.01</versao>
    </DuaServiceHeader>
  </soap12:Header>
  <soap12:Body>
    <duaEmissao xmlns="http://www.sefaz.es.gov.br/duae">
      <duaDadosMsg><![CDATA[{xml_dados_retry}]]></duaDadosMsg>
    </duaEmissao>
  </soap12:Body>
</soap12:Envelope>"""
                    try:
                        debug_req_es2 = os.path.join(os.environ["USERPROFILE"], "Desktop", f"DEBUG_ENVIO_ES_{nfe_num}_R.xml")
                        with open(debug_req_es2, "w", encoding="utf-8") as f:
                            f.write(soap_env_retry)
                    except:
                        pass
                    retorno2 = enviar_soap(soap_env_retry, nfe_num, tag="_R")
                    try:
                        debug_res_es2 = os.path.join(os.environ["USERPROFILE"], "Desktop", f"DEBUG_RESPOSTA_ES_{nfe_num}_R.xml")
                        with open(debug_res_es2, "w", encoding="utf-8") as f:
                            f.write(retorno2)
                    except:
                        pass
                    cStat2_match = re.search(r"(?:cStat&gt;|cStat>|cStat :)(\d+)", retorno2)
                    cStat2 = cStat2_match.group(1) if cStat2_match else ""
                    if cStat2 == "105":
                        match2 = re.search(r"(?:nDua&gt;|nDua>|nDua :)\s*(\d+)", retorno2)
                        nDua2 = match2.group(1) if match2 else "???"
                        sucessos.append(f"NFe {nfe_num} (reenvio) -> DUA {nDua2}")
                        atualizar_protocolo(nfe_num, nDua2)
                    else:
                        xMotivo2_match = re.search(r"(?:xMotivo&gt;|xMotivo>|xMotivo :)\s*([^&<]+)", retorno2)
                        motivo2 = xMotivo2_match.group(1).strip() if xMotivo2_match else ""
                        erros.append(f"NFe {nfe_num} -> reenvio falhou: {motivo2} (Status: {cStat2 if cStat2 else 'N/A'})")
                # fim do novo fluxo
                continue

                # bloco antigo de envio (desativado pelo continue do fluxo novo)
                retorno = ""
                if thumbprint:
                    temp_soap = os.path.join(os.environ['TEMP'], f"dua_es_{nfe_num}.xml")
                    with open(temp_soap, "w", encoding="utf-8") as f: f.write(soap_env)
                    
                    ps_script = f"""
                    $ProgressPreference = 'SilentlyContinue'
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    try {{
                        $body = [System.IO.File]::ReadAllText('{temp_soap}')
                        $cert = Get-ChildItem -Path Cert:\\CurrentUser\\My | Where-Object {{ $_.Thumbprint -eq '{thumbprint.strip()}' }}
                        if (-not $cert) {{ throw "Certificado nao encontrado" }}
                        
                        # Content-Type específico para SOAP 1.2
                        $headers = @{{ "Content-Type" = "application/soap+xml; charset=utf-8" }}
                        $resp = Invoke-WebRequest -Uri "{url}" -Method Post -Body $body -Certificate $cert -Headers $headers -UseBasicParsing
                        Write-Output $resp.Content
                    }} catch {{
                        Write-Output "ERRO: $($_.Exception.Message)"
                        if ($_.Exception.Response) {{
                            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                            Write-Output "DETALHE: $($reader.ReadToEnd())"
                        }}
                        exit 1
                    }}
                    """
                    res = subprocess.run(["powershell", "-Command", ps_script], capture_output=True, text=True)
                    if os.path.exists(temp_soap): os.remove(temp_soap)
                    retorno = res.stdout
                else:
                    from cryptography.hazmat.primitives import serialization
                    import tempfile
                    
                    senha = simpledialog.askstring("Senha", f"Senha do certificado para nota {nfe_num}:", show='*')
                    if not senha: break
                    
                    with open(cert_path, "rb") as f: pfx_data = f.read()
                    p_key, cert, adds = pkcs12.load_key_and_certificates(pfx_data, senha.encode(), default_backend())
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pem") as tmp:
                        tmp.write(p_key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.TraditionalOpenSSL, serialization.NoEncryption()))
                        tmp.write(cert.public_bytes(serialization.Encoding.PEM))
                        pem_path = tmp.name
                    
                    try:
                        resp = requests.post(url, data=soap_env.encode('utf-8'), 
                                           headers={"Content-Type": "application/soap+xml; charset=utf-8"}, 
                                           cert=pem_path, timeout=30)
                        retorno = resp.text
                    finally:
                        if os.path.exists(pem_path): os.remove(pem_path)


                # Salvar debug da resposta ES
                try:
                    debug_path_es = os.path.join(os.environ["USERPROFILE"], "Desktop", f"DEBUG_RESPOSTA_ES_{nfe_num}.xml")
                    with open(debug_path_es, "w", encoding="utf-8") as f:
                        f.write(retorno)
                except:
                    pass

                # Parsing mais robusto
                # Procura por cStat em qualquer formato (escape ou literal, com ou sem prefixo)
                cStat_match = re.search(r"(?:cStat&gt;|cStat>|cStat :)(\d+)", retorno)
                cStat = cStat_match.group(1) if cStat_match else ""

                if cStat == "105":
                    # Sucesso: Extrair número do DUA
                    match = re.search(r"(?:nDua&gt;|nDua>|nDua :)\s*(\d+)", retorno)
                    nDua = match.group(1) if match else "???"
                    sucessos.append(f"NFe {nfe_num} -> DUA {nDua}")
                else:
                    # Erro ou outro status: Extrair motivo
                    # Tenta pegar xMotivo
                    xMotivo_match = re.search(r"(?:xMotivo&gt;|xMotivo>|xMotivo :)\s*([^&<]+)", retorno)
                    motivo = xMotivo_match.group(1).strip() if xMotivo_match else ""
                    
                    if not motivo:
                        # Se não achou xMotivo, tenta pegar a descrição do erro na Resposta SOAP (Fault)
                        fault_match = re.search(r"(?:faultstring&gt;|faultstring>|faultstring :)\s*([^&<]+)", retorno)
                        motivo = fault_match.group(1).strip() if fault_match else "Erro desconhecido na resposta"
                    
                    erros.append(f"NFe {nfe_num} -> {motivo} (Status: {cStat if cStat else 'N/A'})")

            except Exception as e:
                erros.append(f"NFe {nfe_num} -> Erro: {str(e)}")

            except Exception as e:
                erros.append(f"NFe {nfe_num} -> Erro: {str(e)}")

        res_msg = f"RESULTADO ENVIO ES (DUA-E):\n\n✅ Sucessos: {len(sucessos)}\n❌ Erros: {len(erros)}"
        if sucessos: res_msg += "\n\nDETALHES SUCESSOS:\n" + "\n".join(sucessos)
        if erros: res_msg += "\n\nDETALHES ERROS:\n" + "\n".join(erros)
        
        messagebox.showinfo("Retorno Webservice ES", res_msg)

    except Exception as e:
        messagebox.showerror("Erro Crítico ES", f"Falha ao processar lote ES: {e}")

def enviar_lote_webservice():
    """Assina e envia o lote de GNRE selecionado, detectando se é para SEFAZ-PE ou SEFAZ-ES."""
    config = get_configuracoes()
    cert_path = config.get('CERTIFICADO_PATH')
    thumbprint = config.get('CERT_THUMBPRINT')
    
    if not cert_path and not thumbprint:
        messagebox.showwarning("Atenção", "Certificado digital não configurado.\nVá em Configurações > Configurar Pastas.")
        return


    selecionados = tree_gnre.selection()
    if not selecionados:
        # Se nada selecionado, considera todas as notas do grid (comportamento original)
        selecionados = tree_gnre.get_children()
        
    if not selecionados:
        messagebox.showwarning("Atenção", "Nenhuma nota está no grid para enviar.")
        return

    nfe_numeros = [tree_gnre.item(row_id, 'values')[0] for row_id in selecionados]
    
    # Separar notas por UF
    nfe_es = []
    nfe_outros = []
    
    try:
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()
        placeholders = ','.join('?' for _ in nfe_numeros)
        cursor.execute(f"SELECT Nº_NFE, UF_TOMADOR FROM DADOS_GNRE WHERE Nº_NFE IN ({placeholders})", nfe_numeros)
        notas_uf = cursor.fetchall()
        conn.close()
        
        for num, uf in notas_uf:
            if uf == "ES":
                nfe_es.append(num)
            else:
                nfe_outros.append(num)
        
        # Processar Notas do ES via Webservice DUA-E
        if nfe_es:
            if messagebox.askyesno("Confirmar Envio ES", f"Detectamos {len(nfe_es)} notas para o Espírito Santo.\nDeseja enviar via Webservice DUA-E (SEFAZ-ES)?"):
                enviar_dua_es_webservice(nfe_es, thumbprint, cert_path)

        # Processar Notas de outros estados via Webservice GNRE (PE)
        if nfe_outros:
            if messagebox.askyesno("Confirmar Envio Lote", f"Deseja enviar {len(nfe_outros)} notas para o Webservice GNRE (Outros Estados)?"):
                enviar_lote_webservice_pe(nfe_outros, thumbprint, cert_path)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar lote: {e}")

def enviar_lote_webservice_pe(nfe_numeros, thumbprint, cert_path):
    """Lógica original de envio para a SEFAZ-PE (GNRE Nacional)."""
    try:
        # 1. Coletar Dados do Banco
        conn = sqlite3.connect("DADOS_GNRE.db")
        cursor = conn.cursor()
        placeholders = ','.join('?' for _ in nfe_numeros)
        cursor.execute(f"SELECT * FROM DADOS_GNRE WHERE Nº_NFE IN ({placeholders})", nfe_numeros)
        notas_completas = cursor.fetchall()
        cursor.execute("SELECT * FROM DADOS_GNRE LIMIT 1"); emitente = cursor.fetchone()
        conn.close()
        if not emitente: return

        # 2. Gerar XML Lote (Com prefixos gnre: conforme exigência de validação de PE)
        NS_GNRE = "http://www.gnre.pe.gov.br"
        nsmap = {'gnre': NS_GNRE}
        def gnre_tag(tag): return f"{{{NS_GNRE}}}{tag}"
        
        root = lET.Element(gnre_tag("TLote_GNRE"), nsmap=nsmap, versao="2.00")
        guias = lET.SubElement(root, gnre_tag("guias"))
        for nota in notas_completas:
            dg = lET.SubElement(guias, gnre_tag("TDadosGNRE"), versao="2.00")
            lET.SubElement(dg, gnre_tag("ufFavorecida")).text = str(nota[4])
            lET.SubElement(dg, gnre_tag("tipoGnre")).text = "0"
            cem = lET.SubElement(dg, gnre_tag("contribuinteEmitente"))
            id_em = lET.SubElement(cem, gnre_tag("identificacao"))
            lET.SubElement(id_em, gnre_tag("CNPJ")).text = str(emitente[23])
            lET.SubElement(cem, gnre_tag("razaoSocial")).text = str(emitente[24])
            lET.SubElement(cem, gnre_tag("endereco")).text = str(emitente[25])
            lET.SubElement(cem, gnre_tag("municipio")).text = str(emitente[26])
            lET.SubElement(cem, gnre_tag("uf")).text = str(emitente[27])
            lET.SubElement(cem, gnre_tag("cep")).text = str(emitente[28])
            lET.SubElement(cem, gnre_tag("telefone")).text = str(emitente[29])
            
            ignre = lET.SubElement(dg, gnre_tag("itensGNRE"))
            item = lET.SubElement(ignre, gnre_tag("item"))
            lET.SubElement(item, gnre_tag("receita")).text = str(nota[32])
            tipo_orig = "24" if (nota[4] == "PE" and nota[5] == "9") else str(nota[30])
            doc_orig = lET.SubElement(item, gnre_tag("documentoOrigem"), tipo=tipo_orig)
            doc_orig.text = str(nota[31])
            lET.SubElement(item, gnre_tag("produto")).text = "18"
            
            ref = lET.SubElement(item, gnre_tag("referencia"))
            lET.SubElement(ref, gnre_tag("periodo")).text = "0"
            lET.SubElement(ref, gnre_tag("mes")).text = str(nota[16])
            lET.SubElement(ref, gnre_tag("ano")).text = datetime.now().strftime("%Y")
            
            lET.SubElement(item, gnre_tag("dataVencimento")).text = datetime.now().strftime("%Y-%m-%d")
            
            # VALOR GNRE (Tipo 11 - Principal) - MESMO LAYOUT DA FUNÇÃO COMPROVADA
            if nota[4] == "CE":
                v11 = str(nota[6])
            elif nota[5] == "9":
                v11 = str(nota[6])
            elif nota[4] == "AL" or nota[4] == "RJ":
                v11 = str(nota[8])
            else:
                v11 = str(nota[33])
            lET.SubElement(item, gnre_tag("valor"), tipo="11").text = v11
            
            # VALOR GNRE (Tipo 12 - FECP) - MESMO LAYOUT DA FUNÇÃO COMPROVADA
            if nota[4] in ["MG", "PR", "MS", "BA", "CE", "RS", "PE", "MT", "AM"]:
                vfecp = 0.0
            else:
                vfecp = float(nota[7]) + float(nota[39]) if nota[5] == "9" else float(nota[7])
            lET.SubElement(item, gnre_tag("valor"), tipo="12").text = f"{vfecp:.2f}"
            lET.SubElement(item, gnre_tag("convenio")).text = str(nota[12])
            
            cdest = lET.SubElement(item, gnre_tag("contribuinteDestinatario"))
            idest = lET.SubElement(cdest, gnre_tag("identificacao"))
            if nota[4] == "RJ" and nota[5] == "9":
                lET.SubElement(idest, gnre_tag("CNPJ")).text = str(nota[14])
                lET.SubElement(cdest, gnre_tag("razaoSocial")).text = str(nota[3])
                lET.SubElement(cdest, gnre_tag("municipio")).text = str(nota[40][-5:])
            else:
                lET.SubElement(idest, gnre_tag("IE")).text = str(nota[15])
            
            cextras = lET.SubElement(item, gnre_tag("camposExtras") )
            c1 = lET.SubElement(cextras, gnre_tag("campoExtra"))
            lET.SubElement(c1, gnre_tag("codigo")).text = str(nota[21])
            t_v_c1 = lET.SubElement(c1, gnre_tag("valor"))
            if ((nota[4] in ["PR", "MS", "BA"] and nota[5] == "9") or (nota[4] == "AL" and nota[5] == "1")):
                t_v_c1.text = str(nota[9])
            else:
                t_v_c1.text = datetime.now().strftime("%Y-%m-%d")
            
            c2 = lET.SubElement(cextras, gnre_tag("campoExtra"))
            lET.SubElement(c2, gnre_tag("codigo")).text = str(nota[22])
            lET.SubElement(c2, gnre_tag("valor")).text = str(nota[18])
            
            lET.SubElement(dg, gnre_tag("valorGNRE")).text = str(nota[33])
            lET.SubElement(dg, gnre_tag("dataPagamento")).text = datetime.now().strftime("%Y-%m-%d")

        xml_unsigned = lET.tostring(root, encoding='utf-8', xml_declaration=True)
        xml_signed_clean = ""

        # 3. ASSINATURA DIGITAL (TESTE: Removendo para ver se o erro 'Invalid content' some)
        # Em alguns WebServices de Lote 2.0, a assinatura é feita apenas na conexão mTLS.
        xml_signed_clean = xml_unsigned.decode('utf-8')
        import re
        xml_signed_clean = re.sub(r'<\?xml.*?\?>', '', xml_signed_clean, flags=re.IGNORECASE | re.DOTALL).strip()

        # 4. Enviar via SOAP (Seguindo fielmente o WSDL fornecido pelo usuário)
        soap_env = f"""<?xml version="1.0" encoding="utf-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:tns="http://www.gnre.pe.gov.br/webservice/GnreLoteRecepcao">
    <soapenv:Header>
        <tns:gnreCabecMsg>
            <tns:versaoDados>2.00</tns:versaoDados>
        </tns:gnreCabecMsg>
    </soapenv:Header>
    <soapenv:Body>
        <tns:gnreDadosMsg>{xml_signed_clean}</tns:gnreDadosMsg>
    </soapenv:Body>
</soapenv:Envelope>"""
        url = "https://www.gnre.pe.gov.br/gnreWS/services/GnreLoteRecepcao"
        
        # Salva o XML gerado para análise do usuário
        try:
            debug_path = os.path.join(os.environ["USERPROFILE"], "Desktop", "DEBUG_ENVIO_GNRE.xml")
            with open(debug_path, "w", encoding="utf-8") as f:
                f.write(soap_env)
            print(f"DEBUG - XML de envio salvo em: {debug_path}")
        except:
            pass

        print(f"DEBUG - Envelope SOAP: {soap_env[:500]}...")

        if thumbprint:
            # MÉTODO A: Envio via PowerShell para usar o Handshake SSL do Windows (mTLS)
            temp_soap = os.path.join(os.environ['TEMP'], "soap_envelope.xml")
            # Escreve sem BOM para evitar erros de parse no servidor
            with open(temp_soap, "w", encoding="utf-8") as f: f.write(soap_env)
            
            temp_ps1 = os.path.join(os.environ['TEMP'], "send_gnre.ps1")
            ps_script = f"""
            $ProgressPreference = 'SilentlyContinue'
            $WarningPreference = 'SilentlyContinue'
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            try {{
                $body = [System.IO.File]::ReadAllText('{temp_soap}')
                $url = "https://www.gnre.pe.gov.br/gnreWS/services/GnreLoteRecepcao"
                $cert = Get-ChildItem -Path Cert:\\CurrentUser\\My | Where-Object {{ $_.Thumbprint -eq '{thumbprint.strip()}' }}
                if (-not $cert) {{ throw "Certificado '{thumbprint.strip()}' nao encontrado." }}
                
                # SOAPAction completo conforme manual GNRE Nacional
                $headers = @{{ "SOAPAction" = "http://www.gnre.pe.gov.br/webservice/GnreLoteRecepcao/processar" }}
                # Uso do Invoke-WebRequest para garantir o envio correto do header SOAPAction (SOAP 1.1)
                $resp = Invoke-WebRequest -UseBasicParsing -Uri $url -Method Post -Body $body -ContentType "text/xml; charset=utf-8" -Certificate $cert -Headers $headers
                Write-Output $resp.Content
            }} catch {{
                Write-Output "ERRO_SEFAZ: $($_.Exception.Message)"
                if ($_.Exception.Response) {{
                    $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                    Write-Output "DETALHE_ERRO: $($reader.ReadToEnd())"
                }}
                exit 1
            }}
            """
            
            with open(temp_ps1, "w", encoding="utf-8") as f:
                f.write(ps_script)
                
            res = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", temp_ps1], capture_output=True, text=True)
            
            if os.path.exists(temp_soap): os.remove(temp_soap)
            if os.path.exists(temp_ps1): os.remove(temp_ps1)
            
            print(f"DEBUG - Retorno PowerShell (STDOUT): {res.stdout}")
            print(f"DEBUG - Retorno PowerShell (STDERR): {res.stderr}")

            if res.returncode == 0 and res.stdout.strip():
                retorno = res.stdout
                import re
                # ... resto do código de parsing ...
                # Regex robusta para capturar o número do recibo (lote) e a situação
                recibo = re.search(r'<(?:.*:)?(?:numero|numeroRecibo)>(.*?)</(?:.*:)?(?:numero|numeroRecibo)>', retorno)
                situacao = re.search(r'<(?:.*:)?descricao>(.*?)</(?:.*:)?descricao>', retorno)
                
                if recibo or situacao:
                    msg = "RESPOSTA DA SEFAZ-PE:\n\n"
                    if recibo: 
                        num_recibo = recibo.group(1)
                        msg += f"🎫 NÚMERO DO RECIBO (LOTE): {num_recibo}\n"
                        msg += "✅ (Copiado para a área de transferência!)\n"
                        pyperclip.copy(num_recibo)
                        
                        # Salvar Protocolo e Status no Banco de Dados
                        try:
                            conn = sqlite3.connect("DADOS_GNRE.db")
                            cursor = conn.cursor()
                            # Atualiza para todas as notas desse lote
                            placeholders = ','.join('?' for _ in nfe_numeros)
                            status_txt = situacao.group(1) if situacao else "Enviado"
                            cursor.execute(f"UPDATE DADOS_GNRE SET PROTOCOLO_GNRE = ?, STATUS_GNRE = ? WHERE Nº_NFE IN ({placeholders})", 
                                           (num_recibo, status_txt, *nfe_numeros))
                            conn.commit()
                            conn.close()
                            print(f"DEBUG - Protocolo {num_recibo} salvo para as notas: {nfe_numeros}")
                        except Exception as e_db:
                            print(f"Erro ao salvar protocolo no banco: {e_db}")
                    
                    if situacao: 
                        msg += f"📊 SITUAÇÃO: {situacao.group(1)}"
                    
                    messagebox.showinfo("Retorno Webservice", msg)
                    
                    # Perguntar se quer abrir o site de consulta
                    if messagebox.askyesno("Consultar Lote", "Deseja abrir o portal da GNRE para consultar o processamento do lote?"):
                        import webbrowser
                        # Tentamos passar o parâmetro, embora muitos portais GNRE ignorem
                        webbrowser.open(f"https://www.gnre.pe.gov.br/gnre/v/lote/consultar?numeroRecibo={num_recibo}")
                    
                    # Atualizar dashboard e tabelas após sucesso
                    if app.winfo_exists():
                        atualizar_todas_as_tabelas_e_abas()
                else:
                    messagebox.showinfo("Sucesso", "Lote enviado! Verifique o retorno no site da SEFAZ.")
            else:
                # Se falhou, tenta pegar o erro amigável que escrevemos no STDOUT do PowerShell
                erro_msg = res.stdout if res.stdout.strip() else res.stderr
                if not erro_msg: erro_msg = "Erro desconhecido na comunicação."
                messagebox.showerror("Erro no Envio", f"Falha na comunicação com a SEFAZ:\n{erro_msg[:500]}")

        else:
            # MÉTODO B: Envio via Requests (Arquivo .PFX)
            import tempfile
            from cryptography.hazmat.primitives import serialization
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pem") as tmp:
                tmp.write(p_key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.TraditionalOpenSSL, serialization.NoEncryption()))
                tmp.write(cert.public_bytes(serialization.Encoding.PEM))
                pem_path = tmp.name
            
            try:
                resp = requests.post(url, data=soap_env.encode('utf-8'), headers=headers, cert=pem_path, timeout=30)
                os.remove(pem_path)
                if resp.status_code == 200:
                    retorno = resp.text
                    recibo = re.search(r'<(?:.*:)?(?:numero|numeroRecibo)>(.*?)</(?:.*:)?(?:numero|numeroRecibo)>', retorno)
                    situacao = re.search(r'<(?:.*:)?descricao>(.*?)</(?:.*:)?descricao>', retorno)
                    
                    msg = "Lote enviado com sucesso para a SEFAZ-PE!\n\n"
                    if recibo:
                        num_recibo = recibo.group(1)
                        msg += f"🎫 NÚMERO DO RECIBO (LOTE): {num_recibo}\n"
                        msg += "✅ (Copiado para a área de transferência!)\n"
                        pyperclip.copy(num_recibo)

                        # Salvar Protocolo e Status no Banco de Dados (Método Requests)
                        try:
                            conn = sqlite3.connect("DADOS_GNRE.db")
                            cursor = conn.cursor()
                            placeholders = ','.join('?' for _ in nfe_numeros)
                            status_txt = situacao.group(1) if situacao else "Enviado"
                            cursor.execute(f"UPDATE DADOS_GNRE SET PROTOCOLO_GNRE = ?, STATUS_GNRE = ? WHERE Nº_NFE IN ({placeholders})", 
                                           (num_recibo, status_txt, *nfe_numeros))
                            conn.commit()
                            conn.close()
                        except Exception as e_db:
                            print(f"Erro ao salvar protocolo no banco: {e_db}")
                    
                    if situacao: msg += f"📊 STATUS: {situacao.group(1)}"
                    
                    messagebox.showinfo("Retorno SEFAZ-PE", msg)
                    
                    # Perguntar se quer abrir o site de consulta
                    if messagebox.askyesno("Consultar Lote", "Deseja abrir o portal da GNRE para consultar o processamento do lote?"):
                        import webbrowser
                        webbrowser.open(f"https://www.gnre.pe.gov.br/gnre/v/lote/consultar?numeroRecibo={num_recibo}")

                    # Atualizar dashboard e tabelas após sucesso
                    if app.winfo_exists():
                        atualizar_todas_as_tabelas_e_abas()
                else:
                    messagebox.showerror("Erro", f"Falha no envio: {resp.status_code}\n{resp.text[:200]}")
            except Exception as e:
                if os.path.exists(pem_path): os.remove(pem_path)
                raise e

    except Exception as e:
        messagebox.showerror("Erro Crítico PE", f"Erro no Webservice PE: {e}")


ttk.Button(button_frame, text="Gerar GNRE Agrupado", command=gerar_arquivos_gnre_agrupado).grid(row=0, column=0, padx=5)
# Novo Botão Webservice
ttk.Button(button_frame, text="🚀 ENVIAR VIA WEBSERVICE", command=enviar_lote_webservice, style="Success.TButton").grid(row=0, column=8, padx=15)
# Iniciar a aplicação

app.mainloop()