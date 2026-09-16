# -*- coding: utf-8 -*-
"""
Демонстрация конвейера на тестовом комплекте документов.

    python examples/demo.py

Создаёт комплект из трёх файлов, прогоняет извлечение текста, показывает
контроль полноты (один файл намеренно нечитаемый), пишет решение в журнал
и проверяет состояние сторожа.
"""
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pipeline import config as C
from pipeline import completeness, journal, watchdog


def podgotovit_komplekt():
    """Комплект: читаемый docx, читаемый xlsx и 'скан' без текстового слоя."""
    papka = os.path.join(C.DOCS_DIR, "DEMO-001")
    os.makedirs(os.path.join(papka, "_txt"), exist_ok=True)
    try:
        from docx import Document
        d = Document()
        d.add_paragraph("Техническое задание")
        t = d.add_table(rows=2, cols=2)
        t.cell(0, 0).text = "Наименование работ"
        t.cell(0, 1).text = "Объём"
        t.cell(1, 0).text = "Обследование объекта"
        t.cell(1, 1).text = "11,2 км"
        d.save(os.path.join(papka, "tz.docx"))
    except ImportError:
        print("нет python-docx, пропускаю создание .docx")
    # файл, который не даст текста: имитация скана
    io.open(os.path.join(papka, "smeta.pdf"), "wb").write(b"%PDF-1.4 fake scan")
    return papka


def main():
    papka = podgotovit_komplekt()
    print("комплект:", papka, "\n")

    print("--- 1. Извлечение текста")
    from pipeline import extract
    for f in sorted(os.listdir(papka)):
        put = os.path.join(papka, f)
        if f.endswith(".docx"):
            ok = extract.docx_to_text(put, os.path.join(papka, "_txt"))
            print("   %-14s -> %s" % (f, "текст извлечён" if ok else "не удалось"))

    print("\n--- 2. Контроль полноты")
    ne_prochitano = completeness.neprochitannye("DEMO-001")
    if ne_prochitano:
        for imya, prichina in ne_prochitano:
            print("   НЕ ПРОЧИТАНО: %s — %s" % (imya, prichina))
        print("   🔴 КРАСНЫЙ: решение принимать нельзя, комплект неполный")
    else:
        print("   🟢 ЗЕЛЁНЫЙ: весь комплект в тексте")

    print("\n--- 3. Журнал решений")
    nomer = journal.zapisat(
        obekt="DEMO-001",
        reshenie="propuskaem",
        prichina="смета пришла сканом без текстового слоя, объёмы проверить нечем; "
                 "решение по неполному комплекту принимать нельзя",
        prognoz=0)
    print("   записано решение №%d" % nomer)
    journal.zakryt("DEMO-001", itog="подтвердилось", fakt=0)
    for k, v in journal.svodka().items():
        print("   %-26s %s" % (k, v))

    print("\n--- 4. Сторож конвейера")
    print("   до прогона: %s — %s" % watchdog.sostoyanie()[:2])
    watchdog.otmetit_uspeh(naydeno=1)
    print("   после прогона: %s — %s" % watchdog.sostoyanie()[:2])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
