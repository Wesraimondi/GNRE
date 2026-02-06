# GERADOR GNRE SP - REF DINATECNICA ATUALIZAÇÃO 04.11.2024 8:22
import getpass
from io import BytesIO
import os
import sqlite3
import time
import locale
import pyperclip
from tkinter import Menu, simpledialog
#from tkinter.tix import Meter
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import Toplevel, filedialog, messagebox, ttk
from datetime import datetime, timedelta
from PyPDF2 import PdfReader, PdfWriter
#import fitz
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch, cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
import requests
from tabulate import tabulate
import pandas as pd
import win32com

def criar_tabela_emails():
    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS EMAIL_CLIENTES (
            COD_PART TEXT PRIMARY KEY,
            EMAIL TEXT
        )
    """)
    conn.commit()
    conn.close()
criar_tabela_emails()

# --- FUNÇÕES PARA CONFIGURAÇÕES DE CAMINHO ---

def criar_tabela_configuracoes():
    """Cria a tabela para armazenar os caminhos de configuração."""
    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS CONFIGURACOES (
            ID INTEGER PRIMARY KEY,
            PASTA_GNRE_ROOT TEXT,
            PASTA_XML_NFE_ROOT TEXT
        )
    """)
    # Garante que sempre haja um registro para atualizar
    cursor.execute("SELECT COUNT(*) FROM CONFIGURACOES")
    if cursor.fetchone()[0] == 0:
        # Insere valores padrão
        cursor.execute("INSERT INTO CONFIGURACOES (ID, PASTA_GNRE_ROOT, PASTA_XML_NFE_ROOT) VALUES (1, 'D:\\GUIA ST', 'S:\\NFE\\')")
    
    conn.commit()
    conn.close()
    
criar_tabela_configuracoes()

def get_configuracoes():
    """Recupera os caminhos de configuração do banco de dados."""
    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()
    # Assume que a configuração está sempre no ID 1
    cursor.execute("SELECT PASTA_GNRE_ROOT, PASTA_XML_NFE_ROOT FROM CONFIGURACOES WHERE ID = 1")
    resultado = cursor.fetchone()
    conn.close()
    
    # Retorna as configurações ou valores padrão se não encontrar
    if resultado:
        return {
            'PASTA_GNRE_ROOT': resultado[0],
            'PASTA_XML_NFE_ROOT': resultado[1]
        }
    else:
        # Se por algum motivo não encontrar, retorna o padrão
        return {
            'PASTA_GNRE_ROOT': 'D:\\GUIA ST',
            'PASTA_XML_NFE_ROOT': 'S:\\NFE\\'
        }

def salvar_configuracoes(pasta_gnre, pasta_xml_nfe):
    """Salva ou atualiza os caminhos de configuração no banco de dados."""
    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE CONFIGURACOES 
        SET PASTA_GNRE_ROOT = ?, PASTA_XML_NFE_ROOT = ?
        WHERE ID = 1
    """, (pasta_gnre, pasta_xml_nfe))
    conn.commit()
    conn.close()
    messagebox.showinfo("Sucesso", "Configurações de caminho salvas com sucesso!")

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

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

def backup_google_drive():
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)

    # Enviar o arquivo para o Google Drive
    arquivo = drive.CreateFile({"title": "DADOS_GNRE_BACKUP.db"})
    arquivo.SetContentFile("DADOS_GNRE.db")
    arquivo.Upload()

    messagebox.showinfo("Sucesso", "Backup enviado para o Google Drive com sucesso!")


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



def atualizar_tabela_gnre():
    """Atualiza o banco de dados e as tabelas"""
    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM DADOS_GNRE ORDER BY Nº_NFE DESC") 
    linhas = cursor.fetchall()
    for linha in linhas:
        tree.insert("", "end", values=linha)
    conn.close()

def criar_bd():
    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()

     # Tabela de dados principais
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS DADOS_GNRE (Nº_NFE TEXT, DT_EMISSÃO TEXT, COD_PART TEXT, RAZÃO_SOCIAL_TOMADOR TEXT, UF_TOMADOR TEXT, CONTRIBUINTE TEXT, VL_ICMS_UF_DEST TEXT,                
            VL_FECP TEXT, VL_ICMSST TEXT, CHAVE_NFE TEXT, VL_EUA TEXT, DATA_EUA TEXT, PROTOCOLO_ICMS TEXT, PROTOCOLO_OBS TEXT, CNPJ_TOMADOR TEXT, IE TEXT, MÊS TEXT,
            NOME TEXT, OBS_GNRE TEXT, VL_FECP_GNRE_EUA TEXT, PC_CLIENTE TEXT, CAMPO_EXTRA1 TEXT, CAMPO TEXT, CNPJ_EMITENTE TEXT, RAZÃO_SOCIAL_EMITENTE TEXT, ENDEREÇO TEXT,
            COD_MUN TEXT, UF_EMITENTE TEXT, CEP TEXT, TELEFONE TEXT, TIPO TEXT, ORIGEM TEXT, COD_RECEITA, VALOR_TOTAL_GNRE TEXT, CAMINHO_PDF TEXT, CAMINHO_CCE TEXT, CAMINHO_XML TEXT, EMAIL TEXT, RENOMEAR TEXT, vFCPUFDest TEXT, MUNICIPIO TEXT  
        )
    ''')
    conn.commit()
    conn.close()
criar_bd()

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
            
            print(f"PDF {arquivo} associado ao registro {numero_doc}")

        conn.commit()
        messagebox.showinfo("Sucesso", "Caminhos dos PDFs associados ao banco de dados com sucesso.")
        
        atualizar_tabela_gnre()
        

    except Exception as e:
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

            cursor.execute("UPDATE DADOS_GNRE SET CAMINHO_XML = ? WHERE Num_NFE = ?", (caminho_xml, numero_doc))

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
criar_bd()

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
    tree = ET.parse(filepath)
    root = tree.getroot()
    # Namespaces para o XML
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
    # Extraindo dados principais
    nfe_data = {}
    nfe_data['Nº_NFE'] = root.find('.//nfe:ide/nfe:nNF', ns).text if root.find('.//nfe:ide/nfe:nNF', ns) is not None else ""
    nfe_data['DT_EMISSÃO'] = root.find('.//nfe:ide/nfe:dhEmi', ns).text[:10] if root.find('.//nfe:ide/nfe:dhEmi', ns) is not None else ""
    inf_nfe = root.find('.//nfe:infNFe', ns)
    nfe_data['CHAVE_NFE'] = inf_nfe.attrib.get('Id', '')[3:] if inf_nfe is not None else ""
    # Extraindo dados do Tomador
    dest = root.find('.//nfe:dest', ns)
    nfe_data['RAZÃO_SOCIAL_TOMADOR'] = ' '.join(dest.find('nfe:xNome', ns).text.split()[:2]) if dest.find('nfe:xNome', ns) is not None else ""
    nfe_data['UF_TOMADOR'] = dest.find('.//nfe:UF', ns).text if dest is not None and dest.find('.//nfe:UF', ns) is not None else ""
    nfe_data['MUNICIPIO'] = dest.find('.//nfe:cMun', ns).text if dest is not None and dest.find('.//nfe:cMun', ns) is not None else ""
    nfe_data['CONTRIBUINTE'] = dest.find('nfe:indIEDest', ns).text if dest.find('nfe:indIEDest', ns) is not None else ""
    nfe_data['CNPJ_TOMADOR'] = dest.find('nfe:CNPJ', ns).text if dest.find('nfe:CNPJ', ns) is not None else ""
    nfe_data['IE'] = dest.find('nfe:IE', ns).text if dest.find('nfe:IE', ns) is not None else ""
    # Extraindo dados financeiros
    icms_tot = root.find('.//nfe:total/nfe:ICMSTot', ns)
    nfe_data['VL_ICMS_UF_DEST'] = icms_tot.find('nfe:vICMSUFDest', ns).text if icms_tot is not None and icms_tot.find('nfe:vICMSUFDest', ns) is not None else "0"
    nfe_data['VL_FECP'] = icms_tot.find('nfe:vFCP', ns).text if icms_tot is not None and icms_tot.find('nfe:vFCP', ns) is not None else "0"
    nfe_data['VL_ICMSST'] = icms_tot.find('nfe:vST', ns).text if icms_tot is not None and icms_tot.find('nfe:vST', ns) is not None else "0"
    nfe_data['VL_FECP'] = icms_tot.find('nfe:vFCPST', ns).text if icms_tot is not None and icms_tot.find('nfe:vFCP', ns) is not None else "0"
    nfe_data['vFCPUFDest'] = icms_tot.find('nfe:vFCPUFDest', ns).text if icms_tot is not None and icms_tot.find('nfe:vFCPUFDest', ns) is not None else "0"
   
    # Extraindo COD_PART
    inf_adic = root.find('.//nfe:infAdic', ns)
    nfe_data['COD_PART'] = ""  # Inicializa como vazio
    if inf_adic is not None:
        for info in inf_adic.findall('nfe:infAdicional', ns):
            if 'Cliente' in info.text:
                partes = info.text.split("CLIENTE")
                if len(partes) > 1:
                    cliente_info = partes[1].strip().split()
                    nfe_data['COD_PART'] = cliente_info[0] if cliente_info else ""
                break
    # Extraindo infCpl e buscando o número do cliente
    inf_cpl = root.find('.//nfe:infCpl', ns)
    if inf_cpl is not None:
        texto_inf_cpl = inf_cpl.text if inf_cpl.text is not None else ""
        partes = texto_inf_cpl.split("CLIENTE")
        if len(partes) > 1:
            cliente_info = partes[1].strip().split()
            nfe_data['COD_PART'] = cliente_info[0] if cliente_info else nfe_data['COD_PART']
    # Extraindo xPed de det/prod e colocando no campo PC_CLIENTE
    nfe_data['PC_CLIENTE'] = ""  # Inicializa como vazio
    dets = root.findall('.//nfe:det', ns)
    for det in dets:
        prod = det.find('nfe:prod', ns)
        if prod is not None:
            xPed = prod.find('nfe:xPed', ns)
            if xPed is not None and xPed.text:
                nfe_data['PC_CLIENTE'] = xPed.text.strip()
                break  # Para se encontrar o primeiro xPed
    # Extraindo dados do Emitente
    emit = root.find('.//nfe:emit', ns)
    nfe_data['CNPJ_EMITENTE'] = emit.find('nfe:CNPJ', ns).text if emit.find('nfe:CNPJ', ns) is not None else ""
    nfe_data['RAZÃO_SOCIAL_EMITENTE'] = emit.find('nfe:xNome', ns).text if emit.find('nfe:xNome', ns) is not None else ""
    endereco = emit.find('nfe:enderEmit', ns)
    nfe_data['ENDEREÇO'] = f"{endereco.find('nfe:xLgr', ns).text} {endereco.find('nfe:nro', ns).text}" if endereco is not None else ""
    nfe_data['COD_MUN'] = endereco.find('nfe:cMun', ns).text[2:] if endereco is not None and endereco.find('nfe:cMun', ns) is not None else ""
    nfe_data['UF_EMITENTE'] = endereco.find('nfe:UF', ns).text if endereco is not None and endereco.find('nfe:UF', ns) is not None else ""
    nfe_data['CEP'] = endereco.find('nfe:CEP', ns).text if endereco is not None and endereco.find('nfe:CEP', ns) is not None else ""
    nfe_data['TELEFONE'] = endereco.find('nfe:fone', ns).text if endereco is not None and endereco.find('nfe:fone', ns) is not None else ""
    return nfe_data
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
                    ENDEREÇO, COD_MUN, UF_EMITENTE, CEP, TELEFONE, TIPO, ORIGEM, COD_RECEITA, VALOR_TOTAL_GNRE, RENOMEAR, vFCPUFDest, MUNICIPIO)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?,?,?,?,?,?)
            ''', (
                nfe_data['Nº_NFE'], nfe_data['DT_EMISSÃO'], nfe_data['COD_PART'], nfe_data['RAZÃO_SOCIAL_TOMADOR'], nfe_data['UF_TOMADOR'],
                nfe_data['CONTRIBUINTE'], nfe_data['VL_ICMS_UF_DEST'], nfe_data['VL_FECP'],  nfe_data['VL_ICMSST'], nfe_data['CHAVE_NFE'],
                nfe_data['VL_EUA'], nfe_data['DATA_EUA'], nfe_data['PROTOCOLO_ICMS'], nfe_data['PROTOCOLO_OBS'], nfe_data['CNPJ_TOMADOR'],
                nfe_data['IE'], nfe_data['MÊS'], nfe_data['NOME'], nfe_data['OBS_GNRE'], nfe_data['VL_FECP_GNRE_EUA'], nfe_data.get('PC_CLIENTE', ''),
                nfe_data['CAMPO_EXTRA1'], nfe_data['CAMPO'], nfe_data['CNPJ_EMITENTE'], nfe_data['RAZÃO_SOCIAL_EMITENTE'], nfe_data['ENDEREÇO'],
                nfe_data['COD_MUN'], nfe_data['UF_EMITENTE'], nfe_data['CEP'], nfe_data['TELEFONE'],  nfe_data['TIPO'], nfe_data['ORIGEM'],
                nfe_data['COD_RECEITA'], "{:.2f}".format(nfe_data['VALOR_TOTAL_GNRE']), nfe_data['RENOMEAR'], nfe_data['vFCPUFDest'], nfe_data['MUNICIPIO']
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

     
    

    for i, filepath in enumerate(filepaths):
        try:
            nfe_data = extrair_dados_xml(filepath)
            # Verificar se pelo menos uma das tags necessárias está preenchida
            if (float(nfe_data['VL_ICMS_UF_DEST']) > 0 or 
                float(nfe_data['VL_FECP']) > 0 or 
                float(nfe_data['VL_ICMSST']) > 0):
                nfe_data = calcular_valores_adicionais(nfe_data)
                inserir_dados(nfe_data)
    
            
    
            app.update_idletasks()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o arquivo {filepath}: {e}")

    messagebox.showinfo("Sucesso, arquivos processados com sucesso.")
    
    
    

# Exemplo de função para atualizar a tabela na interface Tkinter (se necessário)
def atualizar_tabela_gnre():
    # Sua implementação para atualizar a tabela
    pass

def atualizar_tabela2():
    for row in tree.get_children():
        tree.delete(row)
    conn = sqlite3.connect("DADOS_GNRE.db")
    cursor = conn.cursor()
    # Adicione CANCELADA na consulta, assumindo que é a última coluna (índice 40)
    cursor.execute("SELECT *, CANCELADA FROM DADOS_GNRE ORDER BY Nº_NFE DESC") 
    linhas = cursor.fetchall()
    
    # Configure uma nova tag no inicio da função configurar_interface:
    # tree.tag_configure("cancelada", background="lightgrey") 

    for row in linhas:
        # A nova coluna CANCELADA é a última na tupla 'row'
        cancelada = row[-1] 
        caminho_pdf = row[34] # CAMINHO_PDF é a posição 34 na sua tabela

        tags = ()
        if cancelada == 1:
            tags = ("cancelada",) # Nova tag para cancelada
        elif caminho_pdf:  # PDF associado
            tags = ("anexado",)
        else: # Sem PDF e não cancelada
            tags = ("nao_anexado",)

        # Insere a linha sem a coluna 'CANCELADA' na visualização, se não for necessário na treeview
        tree.insert("", "end", values=row[:-1], tags=tags) 
        
    conn.close()
# ...
def atualizar_todas_as_tabelas_e_abas():
    """Atualiza todas as tabelas e abas do programa"""
    atualizar_tabela2()
    # Adicione outras funções de atualização conforme necessário
print (atualizar_todas_as_tabelas_e_abas)
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

def enviar_email_com_anexo(tree, selecionado):
    """Envia o email com anexo"""
    caminho_pdf = tree.item(selecionado, "values")[34]
    nome_usuario = "Wesley Raimundo" # Substitua "Seu nome aqui" pelo nome do usuário logado
    nfe_numero = tree.item(selecionado, "values")[0]
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = tree.item(selecionado, "values")[37]
    mail.Subject = f"CI GNRE - {nfe_numero}"
    mail.Body = f"Prezados,\n\nSegue em anexo GNRE referente à sua NFe {nfe_numero}.\n\nAtenciosamente,\n{nome_usuario}."
    mail.Attachments.Add(caminho_pdf)
    mail.Send()

def configurar_interface():
    global app,  entry_busca, tree, notebook,  gerar_gnre_tab
    app = Tk()
    app.title("Emissor GNRE 2025 V - 1.0")
    # Centralizar a janela na tela
    #app.iconbitmap(r"C:\Users\Wesley.Raimundo\Desktop\Gerenciamento de Sistemas Wesley\Designer-_3_.ico")
    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()
    x = (screen_width / 2) - (900 / 2)
    y = (screen_height / 2) - (600 / 2)
    app.geometry(f"900x600+{int(x)}+{int(y)}")
    app.resizable(True, True)  # Impedir redimensionamento
    # Alterar cor de fundo da janela
    app.configure(bg="#f0f0f0")  # Cor de fundo cinza claro

    # Aplicar um estilo moderno
    style = ttk.Style()
    style.theme_use("clam")  # Altere para 'alt', 'default', 'classic', etc.
    
    # Estilizar botões
    style.configure("TButton", font=("Arial", 10, "bold"), padding=6, background="#2F4F4F", foreground="white")
    
    # Estilizar labels
    style.configure("TLabel", font=("Arial", 10), background="#f0f0f0")

    # Criar um notebook com abas
    
    notebook = ttk.Notebook(app)
    notebook.pack(fill='both', expand=True)
    def sobre(event=None):
        """Mostra uma janela com informações sobre o programa"""
        sobre_janela = Toplevel(app)
        sobre_janela.title("Sobre")
        sobre_janela.geometry("300x150")
        sobre_janela.resizable(False, False)
        sobre_janela.config(cursor="hand2")  # Mudar o ícone do mouse
        label = ttk.Label(sobre_janela, text="Emissor GNRE 2025 V - 1.0\nDesenvolvido por Wesley Raimundo", anchor='center', font=("Helvetica", 10))
        label.pack(pady=10)
        sobre_janela.protocol("WM_DELETE_WINDOW", sobre_janela.destroy)
        sobre_janela.grab_set()  # Impedir que a janela anterior seja usada enquanto esta estiver aberta

    footer_label = ttk.Label(app, text="Gerenciamento de Sistemas Wesley - WESLEY ROCHA RAIMUNDO", anchor='se', font=("Helvetica", 10))
    footer_label.pack(side='bottom', anchor='se', padx=10, pady=10)
    footer_label.bind("<Button-1>", sobre)

    # Guia para associar PDFs
    main_tab = ttk.Frame(notebook)
    

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
    columns = ["Nº NFE", "Data", "Código", "Cliente"]
    tree = ttk.Treeview(main_tab, columns=columns, show="headings")
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor='center')
    tree.pack(pady=10, fill="both", expand=True)
            # Configura tags para linhas da tabela
# Dentro de configurar_interface, após a criação de 'tree':
# ...
    tree.tag_configure("anexado", background="lightgreen")
    tree.tag_configure("nao_anexado", background="lightcoral")
    tree.tag_configure("cancelada", background="lightgrey") # Adicionar esta linha!
    atualizar_tabela2()
# ...
    
    print (atualizar_tabela_gnre)
    menubar = Menu(app)

    util_menu = Menu(menubar, tearoff=0)
    util_menu.add_command(label="Associar PDFs por Nº NFE", command=associar_pdfs_nf)
    util_menu.add_command(label="Associar XMLs por Pasta Raiz", command=associar_xmls_por_pasta_raiz)
    util_menu.add_command(label="Atualizar BD", command=atualizar_todas_as_tabelas_e_abas)
    #util_menu.add_command(label="Atualizar em Tempo Real", command=atualizar_em_tempo_real)
    util_menu.add_command(label="Atualizar Emails por Código de Parte", command=atualizar_emails_por_cod_part)
    util_menu.add_command(label="Backup na Nuvem", command=backup_google_drive)
    util_menu.add_command(label="Cadastrar E-mail Manual", command=cadastrar_email_cliente)
    util_menu.add_command(label="Criar Backup", command=criar_backup)
    util_menu.add_command(label="Importar E-mails de Arquivo", command=importar_emails_por_arquivo)
    util_menu.add_command(label="Importar XMLs", command=importar_xmls)
    util_menu.add_command(label="Selecionar Pasta de cce_tab", command=selecionar_pasta_cce)
    util_menu.add_command(label="Selecionar Pasta de PDFs", command=selecionar_pasta_pdf)
    util_menu.add_command(label="Selecionar Pasta de XMLs", command=selecionar_pasta_xml)
    menubar.add_cascade(label="Utilitários", menu=util_menu)
    app.config(menu=menubar)
    
    for col in columns:
        tree.heading(col, text=col, anchor='center')
        tree.column(col, width=150, anchor='w')
    tree.pack(pady=5, fill='both', expand=True)
    
            # Criar aba - Gerar GNRE
    gerar_gnre_tab = ttk.Frame(notebook)
    notebook.add(gerar_gnre_tab, text="Gerar GNRE")


configurar_interface()
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
        # INCLUI A CONDIÇÃO: E CANCELADA = 0
        query = f"SELECT * FROM DADOS_GNRE WHERE `Nº_NFE` IN ({placeholders}) AND CANCELADA = 0" 
        cursor.execute(query, numeros)
    else:
        # Comportamento antigo: lista só onde não tem PDF associado E NÃO ESTÁ CANCELADA
        cursor.execute("SELECT * FROM DADOS_GNRE WHERE (CAMINHO_PDF IS NULL OR CAMINHO_PDF = '') AND CANCELADA = 0")

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
        gnre_root = "D:\\GUIA ST" 
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
ttk.Button(button_frame, text="Marcar Cancelada", command=marcar_como_cancelada).grid(row=0, column=4, padx=5) # Novo botão
ttk.Button(button_frame, text="Copiar OBS", command=copiar_obs_gnre).grid(row=0, column=5, padx=5)
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
        gnre_root = "D:\\GUIA ST" # Substituir pela variável de configuração em um refactoring futuro
        
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


ttk.Button(button_frame, text="Gerar GNRE Agrupado", command=gerar_arquivos_gnre_agrupado).grid(row=0, column=0, padx=5)
# Iniciar a aplicação

app.mainloop()