# -*- coding: utf-8 -*-
"""
КОНТРОЛЬ ПОЛНОТЫ: какие документы комплекта НЕ попали в текст и почему.

Главная идея модуля: «текст извлечён» и «документ прочитан» — разные вещи.
Скан без текстового слоя, .doc больше потолка конвертации, сорванный таймаут
LibreOffice, неподдержанный формат — всё это даёт на выходе одинаковую
картину: файла с текстом просто нет. Дальше система считает, что прочитала
комплект, и выдаёт уверенный вывод по половине данных.

Реальный случай: техническое задание на 35 страниц было выложено сканом,
извлечение дало 48 байт, и документ получил оценку «требований нет» —
по одному лишь проекту договора.

Поэтому здесь считается не то, что прочитано, а то, что НЕ прочитано,
с причиной для каждого файла. Пока список не пуст, светофор красный.
"""
import io
import os
import re
import glob

from . import config as C


def neprochitannye(num):
    """Документы закупки, которые НЕ попали в текст, с причиной.

    Скан — не единственный способ остаться непрочитанным. Документ выпадает
    из оценки и когда .doc весит больше потолка конвертации, и когда
    LibreOffice сорвался по таймауту, и когда формат не поддержан. Раньше
    об этом печаталась строка в консоль, и всё: в оценку читаемости попадали
    только PDF-сканы, а пропущенный .doc был неотличим от прочитанного.

    Возвращает список (имя файла, причина). Пустой список = прочитано всё.
    """
    base = os.path.join(C.DOCS_DIR, num)
    if not os.path.isdir(base):
        return []
    txt_dir = os.path.join(base, "_txt")
    gotovye = {os.path.basename(t).lower()
               for t in glob.glob(os.path.join(txt_dir, "*.txt"))}

    def est_tekst(imya):
        stem = os.path.splitext(imya)[0].lower()
        return (stem + ".txt") in gotovye or (imya.lower() + ".txt") in gotovye

    MAX_SRC = 40 * 1024 * 1024
    out, vidno = [], set()
    for f in glob.glob(os.path.join(base, "**", "*"), recursive=True):
        if not os.path.isfile(f) or "_txt" in f or "_png" in f:
            continue
        if not f.lower().endswith((".pdf", ".doc", ".docx", ".xls", ".xlsx", ".rtf")):
            continue
        imya = os.path.basename(f)
        if imya in vidno:
            continue
        vidno.add(imya)
        if est_tekst(imya):
            continue
        razmer = os.path.getsize(f)
        if razmer > MAX_SRC and not f.lower().endswith(".pdf"):
            prichina = "файл %d МБ — больше потолка конвертации" % (razmer // 1048576)
        elif f.lower().endswith(".pdf"):
            prichina = "PDF без текстового слоя (скан)"
        else:
            prichina = "конвертация не дала текста"
        out.append((imya, prichina))
    return out


def scan_check(num, render=6):
    """Ищет в документации PDF-сканы, из которых текст не извлёкся.

    Дыра, найденная 13.08.2026 на закупке Упрдор «Черноморье»: техзадание было
    выложено сканом на 35 страниц, извлечение дало 48 байт, и закупка получила
    оценку «барьеров нет» по одному лишь проекту контракта. Скан молча превращает
    непрочитанное ТЗ в «чисто».

    Возвращает список сканов; первые страницы рендерит в _png/ для чтения глазами.
    """
    import pymupdf
    base = os.path.join(C.DOCS_DIR, num)
    if not os.path.isdir(base):
        return []
    png_dir = os.path.join(base, "_png")
    found, seen = [], set()
    for p in glob.glob(os.path.join(base, "**", "*.pdf"), recursive=True):
        size = os.path.getsize(p)
        if size in seen:                     # тот же файл из соседнего архива
            continue
        seen.add(size)
        try:
            doc = pymupdf.open(p)
        except Exception:
            continue
        probe = min(3, doc.page_count)
        chars = sum(len(doc[i].get_text()) for i in range(probe))
        if chars >= 200 * probe:             # текстовый PDF — читается обычным путём
            continue
        name = os.path.basename(p)
        found.append((name, doc.page_count))
        if render:
            os.makedirs(png_dir, exist_ok=True)
            stem = re.sub(r"[^\w\-]", "_", os.path.splitext(name)[0])[:40]
            for i in range(min(render, doc.page_count)):
                out = os.path.join(png_dir, "%s_p%02d.png" % (stem, i + 1))
                if not os.path.exists(out):
                    doc[i].get_pixmap(dpi=150).save(out)
    return found
