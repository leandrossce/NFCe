# COMO USAR:
#   GUI: python danfe_nfce_pdf.py --gui
#   CLI diretório -> PDFs + Excel: python danfe_nfce_pdf.py "F:\DIRETORIO_XML" "F:\SAIDA" --excel "F:\SAIDA\NFCe_itens.xlsx"
#   CLI arquivo único -> PDF (e opcional Excel): python danfe_nfce_pdf.py "F:\um.xml" "F:\saida.pdf" --excel "F:\itens.xlsx"

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import io
import os
import sys
import math
import argparse
import threading
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime
from pathlib import Path

from lxml import etree as ET  # lxml facilita com namespaces
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import qrcode
from reportlab.lib.utils import ImageReader

# ---- Excel (pandas) ----
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except Exception:
    PANDAS_AVAILABLE = False

# --- GUI (Tkinter) ---
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    TK_AVAILABLE = True
except Exception:
    TK_AVAILABLE = False

NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

# =========================
# Utilitários / Formatação
# =========================

def dec(v):
    if v is None:
        return Decimal("0.00")
    if isinstance(v, Decimal):
        return v
    return Decimal(str(v)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

def br_currency(v):
    # formata 1234.5 -> '1.234,50'
    s = f"{dec(v):,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

def get_text(node, xpath):
    if node is None:
        return ""
    el = node.find(xpath, namespaces=NS)
    return el.text.strip() if el is not None and el.text is not None else ""

def get_dec(node, xpath):
    t = get_text(node, xpath)
    return dec(t) if t else Decimal("0.00")

def wrap_text(c, text, x, y, width, line_height, max_lines=None, fontname="Helvetica", fontsize=9):
    c.setFont(fontname, fontsize)
    words = text.split()
    lines, cur = [], ""
    for w in words:
        test = (cur + " " + w).strip()
        if c.stringWidth(test, fontname, fontsize) <= width:
            cur = test
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    if max_lines is not None:
        lines = lines[:max_lines]
    for i, line in enumerate(lines):
        c.drawString(x, y - i*line_height, line)
    return y - (len(lines) * line_height), len(lines)

# Mapas de pagamento (NTs da NFC-e)
TPAG_MAP = {
    "01": "Dinheiro",
    "02": "Cheque",
    "03": "Cartão de Crédito",
    "04": "Cartão de Débito",
    "05": "Crédito Loja",
    "10": "Vale Alimentação",
    "11": "Vale Refeição",
    "12": "Vale Presente",
    "13": "Vale Combustível",
    "15": "Boleto Bancário",
    "16": "Depósito Bancário",
    "17": "PIX",
    "18": "Transf. bancária / Carteira digital",
    "19": "Fidelidade/Cashback/Crédito Virtual",
    "90": "Outros",
}

def format_chave(ch):
    d = "".join([c for c in ch if c.isdigit()])
    return " ".join([d[i:i+4] for i in range(0, len(d), 4)])

# =========================
# Desenho do DANFE
# =========================

def draw_header(c, emit, ide, dest, chave_acesso, dhEmi_str, page_w, page_h, margin, font_b, font_r):
    y = page_h - margin
    c.setFont(font_b, 11)
    c.drawCentredString(page_w/2, y, "DANFE NFC-e - Documento Auxiliar da Nota Fiscal de Consumidor Eletrônica")
    y -= 6
    c.setLineWidth(0.5)
    c.line(margin, y, page_w - margin, y)
    y -= 8

    # Emitente
    c.setFont(font_b, 10)
    c.drawString(margin, y, get_text(emit, "nfe:xFant") or get_text(emit, "nfe:xNome") or "Emitente")
    y -= 12

    c.setFont(font_r, 9)
    ender = emit.find("nfe:enderEmit", NS) if emit is not None else None
    endereco = []
    if ender is not None:
        endereco.append(f"{get_text(ender, 'nfe:xLgr')}, {get_text(ender, 'nfe:nro')}")
        bairro = get_text(ender, "nfe:xBairro")
        xmun = get_text(ender, "nfe:xMun")
        uf = get_text(ender, "nfe:UF")
        cep = get_text(ender, "nfe:CEP")
        addr2 = " - ".join(filter(None, [bairro, f"{xmun}/{uf}"]))
        if addr2:
            endereco.append(addr2)
        if cep:
            endereco.append(f"CEP {cep}")
    for ln in endereco:
        c.drawString(margin, y, ln)
        y -= 11
    c.drawString(margin, y, f"CNPJ: {get_text(emit,'nfe:CNPJ')}   IE: {get_text(emit,'nfe:IE')}")
    y -= 14

    # Chave e emissão
    c.setFont(font_b, 9)
    c.drawString(margin, y, "CHAVE DE ACESSO:")
    c.setFont(font_r, 9)
    c.drawString(margin+90, y, format_chave(chave_acesso))
    y -= 12

    if dhEmi_str:
        c.setFont(font_b, 9); c.drawString(margin, y, "Emissão:")
        c.setFont(font_r, 9); c.drawString(margin+50, y, dhEmi_str)
    y -= 12

    # Destinatário
    c.setFont(font_b, 9); c.drawString(margin, y, "Consumidor:")
    c.setFont(font_r, 9)
    dest_nome = get_text(dest, "nfe:xNome") or "Não informado"
    dest_doc = get_text(dest, "nfe:CPF") or get_text(dest, "nfe:CNPJ")
    doc_str = f" ({dest_doc})" if dest_doc else ""
    c.drawString(margin+65, y, dest_nome + doc_str)
    y -= 6
    c.line(margin, y, page_w - margin, y)
    return y - 6

def draw_items_header(c, x, y, widths, font_b):
    c.setFont(font_b, 9)
    headers = ["CÓD", "DESCRIÇÃO", "QTD", "UN", "V.UNIT", "V.TOTAL"]
    x0 = x
    for i, h in enumerate(headers):
        c.drawString(x0+2, y-2, h)
        x0 += widths[i]
    y -= 10
    c.setLineWidth(0.3)
    c.line(x, y, x + sum(widths), y)
    return y - 15

def draw_item_row(c, x, y, widths, font_r, prod):
    c.setFont(font_r, 9)
    x0 = x
    col_texts = [
        get_text(prod, "nfe:cProd"),
        get_text(prod, "nfe:xProd"),
        f"{Decimal(get_text(prod,'nfe:qCom') or '0'):,.4f}".replace(",", "X").replace(".", ",").replace("X","."),
        get_text(prod, "nfe:uCom"),
        br_currency(get_text(prod,"nfe:vUnCom")),
        br_currency(get_text(prod,"nfe:vProd")),
    ]
    c.drawString(x0+2, y, col_texts[0][:12]); x0 += widths[0]
    y, used = wrap_text(c, col_texts[1], x0+2, y, widths[1]-4, line_height=10, max_lines=2)
    x0 += widths[1]
    c.drawRightString(x0+25, y + (10*used), col_texts[2]); x0 += widths[2]
    c.drawString(x0+2, y + (10*used), col_texts[3]); x0 += widths[3]
    c.drawRightString(x0+20, y + (10*used), col_texts[4]); x0 += widths[4]
    c.drawRightString(x0+25, y + (10*used), col_texts[5])
    return y - 4

def draw_totals(c, icmstot, y, page_w, margin, font_b, font_r):
    c.setLineWidth(0.3)
    c.line(margin, y, page_w - margin, y)
    y -= 12
    vDesc = get_dec(icmstot, "nfe:vDesc")
    vOutro = get_dec(icmstot, "nfe:vOutro")
    vProd  = get_dec(icmstot, "nfe:vProd")
    vNF    = get_dec(icmstot, "nfe:vNF")
    c.setFont(font_b, 10); c.drawString(margin, y, "Totais")
    y -= 12
    c.setFont(font_r, 9)
    c.drawString(margin, y, f"Valor dos Produtos: {br_currency(vProd)}")
    y -= 12
    c.drawString(margin, y, f"Descontos: {br_currency(vDesc)}    Outros: {br_currency(vOutro)}")
    y -= 14
    c.setFont(font_b, 12)
    c.drawRightString(page_w - margin, y, f"VALOR A PAGAR: {br_currency(vNF)}")
    y -= 10
    return y

def draw_payments(c, pag, y, page_w, margin, font_b, font_r):
    if pag is None:
        return y
    c.setLineWidth(0.3)
    c.line(margin, y, page_w - margin, y)
    y -= 12
    c.setFont(font_b, 10); c.drawString(margin, y, "Pagamentos")
    y -= 12
    c.setFont(font_r, 9)
    dets = pag.findall("nfe:detPag", NS)
    for dp in dets:
        tPag = get_text(dp, "nfe:tPag")
        xPag = get_text(dp, "nfe:xPag")
        vPag = get_dec(dp, "nfe:vPag")
        meio = TPAG_MAP.get(tPag, f"Código {tPag}")
        if xPag:
            meio = f"{meio} ({xPag})"
        c.drawString(margin, y, f"{meio}")
        c.drawRightString(page_w - margin, y, br_currency(vPag))
        y -= 12
    vTroco = get_dec(pag, "nfe:vTroco")
    if vTroco > 0:
        c.setFont(font_b, 9)
        c.drawString(margin, y, "Troco")
        c.drawRightString(page_w - margin, y, br_currency(vTroco))
        y -= 12
    return y

def draw_qrcode_and_footer(c, infNFeSupl, chave, y, page_w, margin, font_r):
    url_qr = get_text(infNFeSupl, "nfe:qrCode") if infNFeSupl is not None else ""
    c.setLineWidth(0.3)
    c.line(margin, y, page_w - margin, y)
    y -= 8
    if url_qr:
        qr_img = qrcode.make(url_qr)
        qr_buf = io.BytesIO()
        qr_img.save(qr_buf, format="PNG")
        qr_buf.seek(0)
        qr_reader = ImageReader(qr_buf)
        size = 34*mm
        c.drawImage(qr_reader, margin, y - size, width=size, height=size, preserveAspectRatio=True, mask='auto')
        c.setFont("Helvetica", 8)
        c.drawString(margin + size + 6, y - 10, "Consulta via leitor de QR Code")
        c.drawString(margin + size + 6, y - 22, "Ou acesse o portal da SEFAZ e informe a chave:")
        c.setFont("Helvetica-Bold", 8)
        c.drawString(margin + size + 6, y - 34, format_chave(chave))
        y -= size + 6
    else:
        c.setFont("Helvetica", 8)
        c.drawString(margin, y, "QR Code não informado no XML.")
        y -= 14

    c.setFont("Helvetica", 7)
    c.drawCentredString(page_w/2, y, "DANFE NFC-e - Não é documento fiscal. Válido como representação simplificada da NFC-e.")
    y -= 10
    return y

# =========================
# Core: PDF e Lote + Excel
# =========================

def robust_extract_chave(root) -> str:
    """
    Tenta extrair a chave de acesso da NFC-e a partir de:
    - infNFe/@Id
    - nfeProc/protNFe/infProt/chNFe
    Retorna apenas dígitos (até 44).
    """
    chave = ""
    nfe = root.find("nfe:NFe", NS)
    if nfe is None and root.tag.endswith("NFe"):
        nfe = root
    if nfe is not None:
        inf = nfe.find("nfe:infNFe", NS)
        if inf is not None:
            chave = (inf.get("Id") or "").replace("NFe", "")
    if not chave:
        ch = root.find(".//nfe:protNFe/nfe:infProt/nfe:chNFe", NS)
        if ch is not None and ch.text:
            chave = ch.text.strip()
    chave = "".join([c for c in chave if c.isdigit()])[:44]
    return chave

def parse_items_for_excel(xml_path: Path):
    """
    Lê um XML e retorna uma lista de dicionários (linhas) com as colunas do Excel.
    """
    rows = []
    tree = ET.parse(str(xml_path))
    root = tree.getroot()

    nfe = root.find("nfe:NFe", NS)
    if nfe is None and root.tag.endswith("NFe"):
        nfe = root
    inf = nfe.find("nfe:infNFe", NS) if nfe is not None else None
    ide  = inf.find("nfe:ide", NS) if inf is not None else None

    dhEmi = get_text(ide, "nfe:dhEmi")
    dhEmi_str = ""
    if dhEmi:
        try:
            dt = datetime.fromisoformat(dhEmi)  # 2025-07-15T15:33:21-03:00
            dhEmi_str = dt.strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            dhEmi_str = dhEmi

    chave = robust_extract_chave(root)

    dets = inf.findall("nfe:det", NS) if inf is not None else []
    for det in dets:
        prod = det.find("nfe:prod", NS)
        if prod is None:
            continue
        qCom = get_text(prod, "nfe:qCom") or "0"
        vUnCom = get_text(prod, "nfe:vUnCom") or "0"
        vProd = get_text(prod, "nfe:vProd") or "0"
        try:
            qCom_f = float(Decimal(qCom))
        except Exception:
            qCom_f = 0.0
        try:
            vUnCom_f = float(Decimal(vUnCom))
        except Exception:
            vUnCom_f = 0.0
        try:
            vProd_f = float(Decimal(vProd))
        except Exception:
            vProd_f = 0.0

        rows.append({
            "DATA EMISSÃO": dhEmi_str,
            "CHAVE ELETRÔNICA": chave,
            "CÓD": get_text(prod, "nfe:cProd"),
            "DESCRIÇÃO": get_text(prod, "nfe:xProd"),
            "QTD": qCom_f,
            "UN": get_text(prod, "nfe:uCom"),
            "V.UNIT": vUnCom_f,
            "V.TOTAL": vProd_f,
        })
    return rows

def export_excel(rows, excel_path: Path, log_fn=None):
    if not rows:
        if log_fn: log_fn("Nenhum item para exportar ao Excel.")
        return
    if not PANDAS_AVAILABLE:
        raise RuntimeError("Exportação para Excel requer pandas. Instale com: pip install pandas openpyxl")
    df = pd.DataFrame(rows, columns=[
        "DATA EMISSÃO","CHAVE ELETRÔNICA","CÓD","DESCRIÇÃO","QTD","UN","V.UNIT","V.TOTAL"
    ])
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(str(excel_path), index=False)
    if log_fn: log_fn(f"[EXCEL] {len(df)} linha(s) exportadas para: {excel_path}")

def make_pdf(xml_path, out_pdf, paper="A4"):
    # Fonte TTF (opcional)
    try:
        pdfmetrics.registerFont(TTFont("DejaVu", "DejaVuSans.ttf"))
        FONT_R = "DejaVu"
        FONT_B = "DejaVu"
    except Exception:
        FONT_R = "Helvetica"
        FONT_B = "Helvetica-Bold"

    tree = ET.parse(xml_path)
    root = tree.getroot()

    nfe = root.find("nfe:NFe", NS)
    if nfe is None and root.tag.endswith("NFe"):
        nfe = root
    inf = nfe.find("nfe:infNFe", NS) if nfe is not None else None

    ide  = inf.find("nfe:ide", NS) if inf is not None else None
    emit = inf.find("nfe:emit", NS) if inf is not None else None
    dest = inf.find("nfe:dest", NS) if inf is not None else None
    total= inf.find("nfe:total/nfe:ICMSTot", NS) if inf is not None else None
    pag  = inf.find("nfe:pag", NS) if inf is not None else None
    infSupl = nfe.find("nfe:infNFeSupl", NS) if nfe is not None else None

    chave = robust_extract_chave(root)
    dhEmi = get_text(ide, "nfe:dhEmi")
    dhEmi_str = ""
    if dhEmi:
        try:
            dt = datetime.fromisoformat(dhEmi)  # 2025-07-15T15:33:21-03:00
            dhEmi_str = dt.strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            dhEmi_str = dhEmi

    # Página
    if str(paper).lower().startswith("80"):
        page_w, page_h = (80*mm, 280*mm)
        margin = 5*mm
    else:
        page_w, page_h = A4
        margin = 12*mm

    c = canvas.Canvas(str(out_pdf), pagesize=(page_w, page_h))

    y = draw_header(c, emit, ide, dest, chave, dhEmi_str, page_w, page_h, margin, FONT_B, FONT_R)

    # Tabela itens
    col_widths = [22*mm, 64*mm if page_w < 100*mm else 90*mm, 10*mm, 14*mm, 25*mm, 28*mm]
    x = margin
    y = draw_items_header(c, x, y, col_widths, FONT_B)

    dets = inf.findall("nfe:det", NS) if inf is not None else []
    for det in dets:
        prod = det.find("nfe:prod", NS)
        row_height = 24  # estimativa
        if y - row_height < 40*mm:
            c.showPage()
            y = page_h - margin
            c.setFont(FONT_B, 11)
            c.drawCentredString(page_w/2, y, "DANFE NFC-e (continuação)")
            y -= 14
            y = draw_items_header(c, x, y, col_widths, FONT_B)
        y = draw_item_row(c, x, y, col_widths, FONT_R, prod)

    # Totais
    y = max(y - 6, 60*mm)
    y = draw_totals(c, total, y, page_w, margin, FONT_B, FONT_R)

    # Pagamentos e troco
    y = draw_payments(c, pag, y, page_w, margin, FONT_B, FONT_R)

    # QRCode + rodapé
    y = draw_qrcode_and_footer(c, infSupl, chave, y, page_w, margin, FONT_R)

    c.showPage()
    c.save()

def extract_chave_from_file(xml_path: Path) -> str:
    try:
        tree = ET.parse(str(xml_path))
        root = tree.getroot()
        return robust_extract_chave(root)
    except Exception:
        return ""

def is_xml_file(p: Path) -> bool:
    return p.is_file() and p.suffix.lower() == ".xml"

def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)

def process_single_xml(xml_path: Path, out_dir: Path, paper: str, force_key_name: bool = True):
    ensure_dir(out_dir)
    chave = extract_chave_from_file(xml_path)
    if not chave:
        # fallback: usa o nome original do arquivo
        stem = xml_path.stem
        out_pdf = out_dir / f"{stem}.pdf"
    else:
        out_pdf = out_dir / f"{chave}.pdf"
    make_pdf(str(xml_path), str(out_pdf), paper=paper)
    return out_pdf

def scan_xmls(in_dir: Path, pattern: str = "*.xml", recursive: bool = False):
    if recursive:
        files = list(in_dir.rglob(pattern))
    else:
        files = list(in_dir.glob(pattern))
    return [p for p in files if is_xml_file(p)]

def process_directory(in_dir: Path, out_dir: Path, paper: str, glob: str, recursive: bool,
                      log_fn=None, progress_fn=None, excel_path: Path | None = None):
    ensure_dir(out_dir)
    xmls = scan_xmls(in_dir, glob or "*.xml", recursive)
    total = len(xmls)
    ok, fail = 0, 0
    rows_accum = []
    if log_fn:
        log_fn(f"Encontrados {total} XML(s) em {in_dir} (padrão: {glob}, recursivo: {recursive})")
    for idx, xp in enumerate(sorted(xmls), start=1):
        try:
            out_pdf = process_single_xml(xp, out_dir, paper, force_key_name=True)
            ok += 1
            if log_fn:
                log_fn(f"[OK] {xp.name} -> {out_pdf.name}")
            # Coleta itens para Excel
            if excel_path is not None:
                rows_accum.extend(parse_items_for_excel(xp))
        except Exception as e:
            fail += 1
            if log_fn:
                log_fn(f"[FALHA] {xp.name}: {e}")
        if progress_fn:
            progress_fn(idx, total)
    # Exporta Excel se solicitado
    if excel_path is not None:
        try:
            export_excel(rows_accum, excel_path, log_fn=log_fn)
        except Exception as ee:
            if log_fn: log_fn(f"[ERRO EXCEL] {ee}")
    if log_fn:
        log_fn(f"[RESUMO] Sucesso: {ok} | Falhas: {fail} | Total: {total}")
    return ok, fail, total

# =========================
# GUI
# =========================

class DanfeGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DANFE NFC-e (XML → PDF)")

        # Vars
        self.var_in_dir = tk.StringVar()
        self.var_out_dir = tk.StringVar()
        self.var_paper = tk.StringVar(value="A4")
        self.var_recursive = tk.BooleanVar(value=False)
        self.var_glob = tk.StringVar(value="*.xml")

        self.var_excel_enable = tk.BooleanVar(value=True)
        self.var_excel_path = tk.StringVar(value="")

        # Layout
        frm = ttk.Frame(self.root, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Linha 1: Origem
        ttk.Label(frm, text="Diretório de origem (XML NFC-e):").grid(row=0, column=0, sticky="w")
        ent_in = ttk.Entry(frm, textvariable=self.var_in_dir, width=60)
        ent_in.grid(row=1, column=0, sticky="ew", padx=(0,6))
        ttk.Button(frm, text="Selecionar…", command=self.pick_in_dir).grid(row=1, column=1, sticky="ew")

        # Linha 2: Saída
        ttk.Label(frm, text="Diretório de saída (PDFs):").grid(row=2, column=0, sticky="w", pady=(8,0))
        ent_out = ttk.Entry(frm, textvariable=self.var_out_dir, width=60)
        ent_out.grid(row=3, column=0, sticky="ew", padx=(0,6))
        ttk.Button(frm, text="Selecionar…", command=self.pick_out_dir).grid(row=3, column=1, sticky="ew")

        # Linha 3: Opções
        opt_frame = ttk.Frame(frm)
        opt_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8,4))
        ttk.Label(opt_frame, text="Tamanho do papel:").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(opt_frame, text="A4", variable=self.var_paper, value="A4").grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(opt_frame, text="80mm (térmica)", variable=self.var_paper, value="80mm").grid(row=0, column=2, sticky="w", padx=(8,0))
        ttk.Checkbutton(opt_frame, text="Buscar recursivamente em subpastas", variable=self.var_recursive).grid(row=0, column=3, sticky="w", padx=(12,0))
        ttk.Label(opt_frame, text="Padrão glob:").grid(row=0, column=4, sticky="e", padx=(12,4))
        ttk.Entry(opt_frame, textvariable=self.var_glob, width=12).grid(row=0, column=5, sticky="w")

        # Linha 4: Excel
        excel_frame = ttk.Frame(frm)
        excel_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(8,4))
        ttk.Checkbutton(excel_frame, text="Salvar Excel com itens", variable=self.var_excel_enable).grid(row=0, column=0, sticky="w")
        ttk.Entry(excel_frame, textvariable=self.var_excel_path, width=48).grid(row=0, column=1, sticky="ew", padx=(8,6))
        ttk.Button(excel_frame, text="Escolher…", command=self.pick_excel).grid(row=0, column=2, sticky="ew")

        # Barra de progresso
        self.progress = ttk.Progressbar(frm, mode="determinate", length=400)
        self.progress.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(8,4))

        # Log
        self.txt = tk.Text(frm, height=12, wrap="word")
        self.txt.grid(row=7, column=0, columnspan=2, sticky="nsew")
        frm.rowconfigure(7, weight=1)

        # Botões
        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=8, column=0, columnspan=2, pady=(8,0), sticky="e")
        ttk.Button(btn_frame, text="Converter", command=self.start_conversion).grid(row=0, column=0, padx=(0,6))
        ttk.Button(btn_frame, text="Sair", command=self.root.destroy).grid(row=0, column=1)

    def pick_in_dir(self):
        d = filedialog.askdirectory(title="Escolha o diretório com XML da NFC-e")
        if d:
            self.var_in_dir.set(d)

    def pick_out_dir(self):
        d = filedialog.askdirectory(title="Escolha o diretório de saída para PDFs")
        if d:
            self.var_out_dir.set(d)

    def pick_excel(self):
        f = filedialog.asksaveasfilename(
            title="Salvar Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if f:
            self.var_excel_path.set(f)

    def log(self, msg: str):
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.root.update_idletasks()

    def set_progress(self, current, total):
        if total <= 0:
            self.progress["value"] = 0
            self.progress["maximum"] = 1
            return
        self.progress["maximum"] = total
        self.progress["value"] = current
        self.root.update_idletasks()

    def start_conversion(self):
        in_dir = Path(self.var_in_dir.get().strip())
        out_dir = Path(self.var_out_dir.get().strip())
        paper = self.var_paper.get()
        recursive = bool(self.var_recursive.get())
        glob = self.var_glob.get().strip() or "*.xml"

        # Excel path (default se vazio)
        excel_path = None
        if bool(self.var_excel_enable.get()):
            excel_gui = self.var_excel_path.get().strip()
            if excel_gui:
                excel_path = Path(excel_gui)
            else:
                # padrão: saida/NFCe_itens.xlsx
                if out_dir:
                    excel_path = out_dir / "NFCe_itens.xlsx"

        if not in_dir.exists() or not in_dir.is_dir():
            messagebox.showerror("Erro", "Selecione um diretório de origem válido.")
            return
        if not out_dir.exists():
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível criar o diretório de saída:\n{e}")
                return

        # roda em thread para não travar a GUI
        th = threading.Thread(target=self._run_conversion, args=(in_dir, out_dir, paper, glob, recursive, excel_path), daemon=True)
        th.start()

    def _run_conversion(self, in_dir: Path, out_dir: Path, paper: str, glob: str, recursive: bool, excel_path: Path | None):
        # reset UI
        self.txt.delete("1.0", "end")
        self.set_progress(0, 1)
        try:
            def log_fn(m): self.log(m)
            def progress_fn(cur, tot): self.set_progress(cur, tot)

            ok, fail, total = process_directory(
                in_dir, out_dir, paper=paper, glob=glob, recursive=recursive,
                log_fn=log_fn, progress_fn=progress_fn, excel_path=excel_path
            )
            msg = f"Processo concluído. Sucesso: {ok} | Falhas: {fail} | Total: {total}"
            self.log(msg)
            messagebox.showinfo("Concluído", msg)
        except Exception as e:
            self.log(f"[ERRO] {e}")
            messagebox.showerror("Erro", str(e))

    def run(self):
        self.root.mainloop()

# =========================
# CLI
# =========================

def main():
    ap = argparse.ArgumentParser(
        description="Gerar DANFE NFC-e (PDF) a partir de XML (arquivo, diretório) ou via GUI; opcionalmente exporta itens para Excel."
    )
    ap.add_argument("entrada", nargs="?", help="Caminho do XML OU diretório contendo XMLs (opcional se usar --gui)")
    ap.add_argument("saida", nargs="?", help="Caminho do PDF (se entrada for arquivo) OU diretório de saída (se entrada for diretório)")
    ap.add_argument("--paper", default="A4", help="A4 ou 80mm (padrão: A4)")
    ap.add_argument("--glob", default="*.xml", help="Padrão de busca quando entrada é diretório (padrão: *.xml)")
    ap.add_argument("--recursive", action="store_true", help="Buscar recursivamente em subpastas quando entrada é diretório")
    ap.add_argument("--use-chave", action="store_true", help="(CLI) Nomear PDFs pela chave de acesso (se disponível)")
    ap.add_argument("--excel", help="Caminho do Excel de itens. Se omitido e 'saida' for diretório, salva em SAIDA/NFCe_itens.xlsx")
    ap.add_argument("--gui", action="store_true", help="Abrir interface gráfica")

    args = ap.parse_args()

    # Se GUI foi pedida (ou nenhum argumento passado), abre GUI
    if args.gui or (args.entrada is None and args.saida is None):
        if not TK_AVAILABLE:
            print("[ERRO] Tkinter não está disponível neste ambiente.", file=sys.stderr)
            sys.exit(2)
        app = DanfeGUI()
        app.run()
        return

    entrada = Path(args.entrada) if args.entrada else None
    saida = Path(args.saida) if args.saida else None

    if entrada is None or saida is None:
        print("Uso (CLI): python danfe_nfce_pdf.py <entrada> <saida> [--paper A4|80mm] [--glob '*.xml'] [--recursive] [--use-chave] [--excel caminho.xlsx]", file=sys.stderr)
        sys.exit(2)

    # Resolve caminho Excel padrão quando aplicável
    excel_path = None
    if args.excel:
        excel_path = Path(args.excel)
    elif entrada.is_dir() and (not saida.suffix.lower() == ".pdf"):
        # padrão: SAIDA/NFCe_itens.xlsx
        excel_path = saida / "NFCe_itens.xlsx"

    # CLI: arquivo único ou diretório
    if entrada.is_dir():
        if saida.suffix.lower() == ".pdf":
            print("[ERRO] Para entrada em diretório, 'saida' deve ser um diretório (e não .pdf).", file=sys.stderr)
            sys.exit(2)
        process_directory(
            entrada, saida, paper=args.paper, glob=args.glob, recursive=args.recursive,
            excel_path=excel_path
        )
    elif entrada.is_file():
        # Se saída for arquivo .pdf, gera PDF com esse nome; Excel (se solicitado) terá apenas os itens desse XML.
        if saida.suffix.lower() == ".pdf":
            make_pdf(str(entrada), str(saida), paper=args.paper)
            print(f"OK: PDF gerado em {saida}")
            if excel_path is not None:
                try:
                    rows = parse_items_for_excel(entrada)
                    export_excel(rows, excel_path)
                except Exception as ee:
                    print(f"[ERRO EXCEL] {ee}", file=sys.stderr)
        else:
            ensure_dir(saida)
            if args.use_chave:
                out_pdf = saida / f"{extract_chave_from_file(entrada) or entrada.stem}.pdf"
            else:
                out_pdf = saida / f"{entrada.stem}.pdf"
            make_pdf(str(entrada), str(out_pdf), paper=args.paper)
            print(f"OK: PDF gerado em {out_pdf}")
            if excel_path is not None:
                try:
                    rows = parse_items_for_excel(entrada)
                    export_excel(rows, excel_path)
                except Exception as ee:
                    print(f"[ERRO EXCEL] {ee}", file=sys.stderr)
    else:
        print(f"[ERRO] Caminho de entrada inválido: {entrada}", file=sys.stderr)
        sys.exit(2)

if __name__ == "__main__":
    main()
