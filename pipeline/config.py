# -*- coding: utf-8 -*-
"""Пути и настройки. Всё, что зависит от машины, — только здесь."""
import os
import shutil

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Папка с комплектами документов: DOCS_DIR/<идентификатор комплекта>/файлы
DOCS_DIR = os.environ.get("DOC_PIPELINE_DOCS", os.path.join(BASE, "data", "docs"))

# База журнала решений
DB = os.environ.get("DOC_PIPELINE_DB", os.path.join(BASE, "data", "journal.db"))

# LibreOffice нужен для .doc, .rtf и старых .xls. Для .docx и .xlsx не нужен —
# они читаются своим кодом, см. extract.py
SOFFICE = os.environ.get("SOFFICE") or shutil.which("soffice") or shutil.which("libreoffice") or "soffice"

# Метка последнего успешного прогона: по ней сторож понимает, жив ли конвейер.
# Намеренно НЕ mtime базы: базу трогают и другие процессы, из-за чего она
# «выглядит свежей», и предупреждение опаздывает на несколько суток.
STAMP = os.path.join(BASE, "data", "last_run.stamp")
