# -*- coding: utf-8 -*-
"""
ИЗВЛЕЧЕНИЕ ТЕКСТА ИЗ ДОКУМЕНТОВ: .docx, .xlsx, .xls, .doc, .rtf

Зачем свой код, если есть LibreOffice. Потому что конвертация «как все»
молча теряет данные, а потерянные данные выглядят точно так же, как
прочитанные: файл есть, текст есть, решение принято — по пустому месту.

Три реальных случая, из-за которых написан этот модуль:

1. **LibreOffice теряет таблицы .docx.** На техническом задании закупки
   конвертация дала 712 символов вместо 30 262: всё задание было набрано
   таблицей 13x3. Система оценила документ по титульному листу.
2. **Excel не читался вообще.** LibreOffice вызывался с фильтром Writer,
   Calc на нём падает с кодом 1 и файл на выходе просто не появляется.
   Молча. Полгода обоснования цен и ведомости объёмов не попадали в разбор.
3. **Колонтитулы и сноски не отдаёт даже python-docx.** А там лежит
   содержательное: в одном документе в колонтитуле оказались ФИО и телефон
   ответственного лица, в сносках — размер штрафа, которого в тексте не было.

Вывод, который стоит за модулем: недостающие данные должны падать громко.
Проверка полноты — в `completeness.py`.
"""
import io
import os
import re
import glob
import shutil
import subprocess

from . import config as C


def docx_to_text(src, outdir):
    """Текст из .docx СВОИМИ силами — с абзацами И таблицами.

    24.08.2026: конвертация через LibreOffice в txt теряет содержимое таблиц.
    На ТЗ закупки 0160300042626000054 (КСОДД Пугачёвского района) это дало
    712 символов вместо 30 262: всё задание было набрано таблицей 13×3.
    Радар оценил закупку по титульному листу и написал «барьеров нет» —
    оценка по пустому тексту выглядит так же, как оценка по прочитанному.

    В госзакупках ТЗ таблицей — обычное дело, поэтому .docx читаем сами,
    а LibreOffice оставляем для .doc, .rtf и таблиц Excel.
    """
    try:
        from docx import Document
    except ImportError:
        return False
    try:
        d = Document(src)
    except Exception:
        return False
    parts = [par.text for par in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            cells, seen = [], set()
            for c in row.cells:
                if id(c._tc) in seen:      # объединённые ячейки не дублируем
                    continue
                seen.add(id(c._tc))
                cells.append(c.text.strip())
            line = chr(9).join(x for x in cells if x)
            if line:
                parts.append(line)
    # Колонтитулы, сноски и текстовые врезки python-docx не отдаёт вовсе,
    # а там лежит содержательное. Проверено 12.09.2026 на doc_03.docx закупки
    # по защите информации: в колонтитуле — ФИО и телефон ответственного лица
    # заказчика (готовый контакт для обзвона), в сносках — требование
    # «состав объектов информатизации приведён в Таблице № 1». Теряли 5 %
    # текста, но не случайных, а именно тех, что меняют решение.
    hvosty = []
    try:
        import zipfile
        with zipfile.ZipFile(src) as z:
            for imya in sorted(z.namelist()):
                if not (imya.startswith("word/") and imya.endswith(".xml")):
                    continue
                if imya == "word/document.xml":
                    continue
                if not re.match(r"word/(header|footer|footnotes|endnotes|comments)",
                                imya):
                    continue
                xml = z.read(imya).decode("utf-8", "replace")
                kuski = re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml)
                kusok = " ".join(x.strip() for x in kuski if x.strip())
                if len(kusok) > 30:            # номера страниц не нужны
                    hvosty.append(kusok)
    except Exception:
        pass
    if hvosty:
        vidno = set()
        unikalnye = [h for h in hvosty if not (h in vidno or vidno.add(h))]
        parts.append("### Колонтитулы, сноски и врезки")
        parts.extend(unikalnye)

    text = chr(10).join(parts)
    if not text.strip():
        return False
    out = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".txt")
    io.open(out, "w", encoding="utf-8").write(text)
    return True


def xlsx_to_text(src, outdir):
    """Excel в текст через openpyxl — все листы, а не первый.

    🔴 Найдено 12.09.2026: ни один .xlsx и .xls в проекте никогда не читался.
    LibreOffice вызывался с фильтром «txt:Text (encoded):UTF8», а это фильтр
    Writer — Calc на нём падает с «Please verify input parameters», код 1,
    и файл на выходе не появляется. Молча: в _txt просто не было файла.
    Цена — обоснование НМЦК, ведомости объёмов и сметы не попадали ни в оценку,
    ни в разбор ТЗ, хотя именно там лежат объёмы работ.

    Формат вывода: заголовок листа, затем строки с ячейками через табуляцию.
    Пустые строки выбрасываются, иначе текст раздувается пустотой.
    """
    try:
        import openpyxl
    except ImportError:
        return False
    try:
        wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    except Exception:
        return False
    parts = []
    for ws in wb.worksheets:
        parts.append("### Лист: %s" % ws.title)
        for row in ws.iter_rows(values_only=True):
            cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
            if cells:
                parts.append(chr(9).join(cells))
    try:
        wb.close()
    except Exception:
        pass
    text = chr(10).join(parts)
    if len(text.strip()) < 20:
        return False
    out = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".txt")
    io.open(out, "w", encoding="utf-8").write(text)
    return True


def xls_to_text(src, outdir):
    """Старый .xls: сначала LibreOffice переводит в .xlsx, дальше openpyxl.

    Прямой перевод .xls в csv отдаёт только первый лист, а в обосновании НМЦК
    листов обычно несколько.
    """
    import tempfile
    tmp = tempfile.mkdtemp()
    try:
        r = subprocess.run([C.SOFFICE, "--headless", "--convert-to", "xlsx",
                            "--outdir", tmp, src], capture_output=True, timeout=180)
        if r.returncode != 0:
            return False
        got = glob.glob(os.path.join(tmp, "*.xlsx"))
        if not got:
            return False
        # имя берём от исходника, чтобы .txt лёг рядом с остальными
        ok = xlsx_to_text(got[0], outdir)
        if ok:
            staryy = os.path.join(outdir, os.path.splitext(os.path.basename(got[0]))[0] + ".txt")
            nuzhnyy = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".txt")
            if staryy != nuzhnyy and os.path.exists(staryy):
                try:
                    os.replace(staryy, nuzhnyy)
                except OSError:
                    pass
        return ok
    except Exception:
        return False
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def doc_to_text(src, outdir, cap_mb=200, limit_s=120):
    """Старый .doc — в текст ЧЕРЕЗ промежуточный .docx, а не напрямую.

    13.09.2026. Для .docx таблицы мы читаем сами (см. docx_to_text), а .doc
    отдавался LibreOffice сразу в txt — и там повторялась ровно та беда,
    ради которой docx_to_text и писался: содержимое таблиц пропадает.
    Поймано на порядке оценки заявок закупки 0167300000526000636 (Тюмень,
    doc_05.doc): между заголовками «Раздел II. Критерии и показатели оценки»
    и «Раздел III» в тексте ПУСТО, а в самом документе там таблица с весами
    «цена 30 %, квалификация 70 %» — то есть единственная цифра, по которой
    решается, можем мы выиграть конкурс или нет.

    Двухходовка: LibreOffice переводит .doc в .docx (таблицы сохраняются
    как таблицы), дальше работает наш docx_to_text. Имя промежуточного
    файла совпадает с исходным, поэтому текст ложится в <имя>.txt, как
    и раньше. Промежуточный .docx удаляется — он нужен только транзитом.
    """
    # Пути только абсолютные: LibreOffice трактует относительный --outdir
    # от СВОЕГО рабочего каталога, и файл уходит мимо. 13.09.2026 на этом
    # переконвертация договора закупки 32616369754 молча вернула «сбой»,
    # хотя сама конвертация проходила успешно.
    src = os.path.abspath(src)
    outdir = os.path.abspath(outdir)
    tmp = os.path.join(outdir, "_conv")
    os.makedirs(tmp, exist_ok=True)
    try:
        p = subprocess.Popen(
            [C.SOFFICE, "--headless", "--convert-to", "docx",
             "--outdir", tmp, src],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except Exception:
        return False
    started = time.time()
    while p.poll() is None:
        if time.time() - started > limit_s:
            _kill(p)
            print("   ~ перевод .doc в .docx прерван по времени: %s"
                  % os.path.basename(src))
            return False
        grown = sum(os.path.getsize(t) for t in glob.glob(os.path.join(tmp, "*"))
                    if os.path.isfile(t))
        if grown > cap_mb * 1048576:
            _kill(p)
            print("   ~ перевод .doc в .docx прерван: временные файлы выросли "
                  "до %d МБ на %s" % (grown // 1048576, os.path.basename(src)))
            return False
        time.sleep(0.5)
    poluchen = os.path.join(tmp, os.path.splitext(os.path.basename(src))[0] + ".docx")
    if not os.path.exists(poluchen):
        # Имя могло измениться (пробелы, кириллица) — берём единственный .docx
        est = glob.glob(os.path.join(tmp, "*.docx"))
        if len(est) != 1:
            return False
        poluchen = est[0]
    ok = docx_to_text(poluchen, outdir)
    try:
        os.remove(poluchen)
    except OSError:
        pass
    return ok


def load_text(num):
    parts = []
    for t in glob.glob(os.path.join(C.DOCS_DIR, num, "_txt", "*.txt")):
        try:
            parts.append(io.open(t, encoding="utf-8", errors="ignore").read())
        except Exception:
            pass
    return "\n".join(parts)


# ──────────────────────────── 3. ОЦЕНКА ────────────────────────────
