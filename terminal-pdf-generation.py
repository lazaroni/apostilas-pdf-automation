#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import io
import os
import re
import sys
import winsound
from pathlib import Path

from openpyxl import load_workbook
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader, PdfWriter
from pypdf.generic import RectangleObject, NameObject
from tqdm import tqdm

# =========================
# Configurações ajustáveis
# =========================

FONT_SIZE_PT = 12
TEXT_COLOR = Color(1, 0, 0)
STAMP_ALL_PAGES = True
OUTPUT_DIR_NAME = "saida_pdfs"

ARIAL_CANDIDATES = [
    "Arial.ttf",
    "arial.ttf",
    "C:/Windows/Fonts/arial.ttf",
    "/Library/Fonts/Arial.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
    "/usr/share/fonts/truetype/msttcorefonts/arial.ttf",
]

# =========================
# Funções utilitárias
# =========================

def cm_to_pt(cm: float) -> float:
    return (cm / 2.54) * 72.0

def register_arial_or_fallback() -> str:
    for candidate in ARIAL_CANDIDATES:
        if os.path.isfile(candidate):
            try:
                pdfmetrics.registerFont(TTFont("Arial", candidate))
                print(f"[Fonte] Arial registrada com sucesso: {candidate}")
                return "Arial"
            except Exception:
                pass
    print("[Fonte] Arial não encontrada. Usando Helvetica como fallback.")
    return "Helvetica"

def normalize_cpf(value) -> str | None:
    if value is None:
        return None
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = re.sub(r"\D", "", s)
    return digits if digits else None

def sanitize_filename(name: str) -> str:
    return re.sub(r'[\/\\\:\*\?\"\<\>\|]', "_", name).strip()

def make_overlay(page_width: float, page_height: float, text: str, font_name: str) -> PdfReader:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_width, page_height))
    c.setFillColor(TEXT_COLOR)
    c.setFont(font_name, FONT_SIZE_PT)
    x = cm_to_pt(1.5)
    y = page_height - cm_to_pt(0.5) - FONT_SIZE_PT
    c.drawString(x, y, text)
    c.save()
    buf.seek(0)
    return PdfReader(buf)

def normalizar_pdf_temporario(entrada: Path) -> Path:
    print(f"[PDF] Normalizando PDF: {entrada.name}")
    leitor = PdfReader(str(entrada))
    gravador = PdfWriter()

    for pagina in leitor.pages:
        obj = pagina.get_object()
        if "/MediaBox" not in obj:
            obj[NameObject("/MediaBox")] = RectangleObject([0, 0, 595, 842])
        gravador.add_page(pagina)

    gravador.remove_links()
    gravador.add_metadata({})
    gravador.encrypt("")

    saida = entrada.parent / "pdf_normalizado_temp.pdf"
    with open(saida, "wb") as f:
        gravador.write(f)

    print(f"[PDF] PDF normalizado salvo como: {saida.name}")
    return saida

def stamp_pdf_for_student(src_pdf_path: Path, out_pdf_path: Path, text_line: str, font_name: str):
    reader = PdfReader(str(src_pdf_path))
    writer = PdfWriter()

    page_bar = tqdm(total=len(reader.pages), desc="📄 Páginas", unit="página", leave=True)

    for i, page in enumerate(reader.pages):
        if (i == 0) or STAMP_ALL_PAGES:
            media_box = page.mediabox
            pw = float(media_box.width)
            ph = float(media_box.height)
            overlay_reader = make_overlay(pw, ph, text_line, font_name)
            overlay_page = overlay_reader.pages[0]
            page.merge_page(overlay_page)

        writer.add_page(page)
        page_bar.update(1)

    page_bar.close()

    writer.add_metadata({
        "/Producer": "Python pypdf",
        "/Author": "Automação",
        "/Title": text_line
    })

    out_pdf_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_pdf_path, "wb") as f:
        writer.write(f)

    print(f"[✔] PDF salvo: {out_pdf_path.name}")

def read_students_from_xlsx(xlsx_path: Path) -> list[tuple[str, str]]:
    print(f"[Planilha] Lendo dados de: {xlsx_path.name}")
    wb = load_workbook(filename=str(xlsx_path), data_only=True)
    ws = wb.active
    students: list[tuple[str, str]] = []

    for row in ws.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        name_val, cpf_val = row
        if name_val is None:
            continue
        name = str(name_val).strip()
        cpf = normalize_cpf(cpf_val)
        if not name or not cpf:
            continue
        students.append((name.upper(), cpf))

    print(f"[Planilha] {len(students)} alunos encontrados.")
    return students

def tocar_som_assincrono(nome_arquivo):
    try:
        winsound.PlaySound(nome_arquivo, winsound.SND_FILENAME | winsound.SND_ASYNC)
    except Exception as e:
        print(f"⚠️ Erro ao tocar {nome_arquivo}: {e}")

# =========================
# Função principal
# =========================

def main():
    print("📄 Iniciando processo de geração de PDFs personalizados...\n")

    if len(sys.argv) >= 3:
        pdf_original = Path(sys.argv[1]).expanduser().resolve()
        xlsx_path = Path(sys.argv[2]).expanduser().resolve()
    else:
        pdf_in = input("Caminho do PDF-base: ").strip().strip('"').strip("'")
        xlsx_in = input("Caminho da planilha XLSX: ").strip().strip('"').strip("'")
        pdf_original = Path(pdf_in).expanduser().resolve()
        xlsx_path = Path(xlsx_in).expanduser().resolve()

    if not pdf_original.exists():
        print(f"[Erro] Arquivo PDF não encontrado: {pdf_original}")
        sys.exit(1)
    if not xlsx_path.exists():
        print(f"[Erro] Planilha XLSX não encontrada: {xlsx_path}")
        sys.exit(1)

    resposta = input('Você já possui um arquivo PDF compatível? Digite "s" para sim e "n" para não: ').strip().lower()
    if resposta == "s":
        pdf_path = pdf_original
        print(f"[PDF] Usando PDF original sem normalização: {pdf_path.name}")
    elif resposta == "n":
        pdf_path = normalizar_pdf_temporario(pdf_original)
    else:
        print("[Erro] Resposta inválida. Digite apenas 's' ou 'n'.")
        sys.exit(1)

    students = read_students_from_xlsx(xlsx_path)
    if not students:
        print("[Erro] Nenhum aluno válido encontrado na planilha.")
        sys.exit(1)

    font_name = register_arial_or_fallback()
    out_dir = pdf_original.parent / OUTPUT_DIR_NAME
    out_dir.mkdir(exist_ok=True)

    print(f"\n📦 Gerando PDFs para {len(students)} alunos...\n")

    for i, (name, cpf) in enumerate(students):
        text_line = f"{name} | CPF: {cpf}"
        file_name = sanitize_filename(f"{name}_{cpf}") + ".pdf"
        out_path = out_dir / file_name
        stamp_pdf_for_student(pdf_path, out_path, text_line, font_name)

        if i < len(students) - 1:
            tocar_som_assincrono("whatsapp_message.wav")

    # Toca som final antes de abrir o diretório
    tocar_som_assincrono("wow.wav")

    if resposta == "n":
        try:
            pdf_path.unlink()
            print(f"\n🧹 PDF temporário removido: {pdf_path.name}")
        except Exception:
            print(f"\n⚠️ Não foi possível remover o PDF temporário: {pdf_path.name}")

    try:
        os.startfile(out_dir)
    except Exception as e:
        print(f"\n⚠️ Erro ao abrir o diretório: {e}")

    print(f"\n✅ Concluído. Arquivos gerados em: {out_dir}\n")

if __name__ == "__main__":
    main()
