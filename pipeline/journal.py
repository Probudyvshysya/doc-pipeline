# -*- coding: utf-8 -*-
"""
ЖУРНАЛ РЕШЕНИЙ: что система предложила, что человек решил, чем это кончилось.

Зачем. Автоматика каждый день выдаёт рекомендации, человек по ним принимает
решения — и на этом цепочка обрывается. Нигде не записано, что было взято
в работу, почему, и совпал ли прогноз с фактом. Без этого система не учится:
одни и те же ошибки повторяются, а качество рекомендаций нечем измерить.

Здесь решение фиксируется в момент принятия — с причиной и датой, — а позже
дописывается фактический исход. Раз в месяц отдельный проход сверяет прогнозы
с фактами и показывает, где система систематически ошибается.

Ключевое требование к полю «причина»: это не метка, а текст, который через
полгода объяснит постороннему, почему решили именно так. Записи вида
«не подходит» бесполезны: они не позволяют понять, изменились обстоятельства
или ошиблось суждение.
"""
import io
import os
import sqlite3
import datetime

from . import config as C

SHEMA = """
CREATE TABLE IF NOT EXISTS resheniya (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    data        TEXT NOT NULL,
    obekt       TEXT NOT NULL,
    reshenie    TEXT NOT NULL,      -- beryom | propuskaem | dumaem
    prichina    TEXT NOT NULL,
    prognoz     REAL,               -- ожидаемый результат в деньгах
    itog        TEXT,               -- фактический исход, дописывается позже
    fakt        REAL,               -- фактический результат в деньгах
    data_itoga  TEXT
);
"""


def _db():
    os.makedirs(os.path.dirname(C.DB), exist_ok=True)
    c = sqlite3.connect(C.DB)
    c.executescript(SHEMA)
    return c


def zapisat(obekt, reshenie, prichina, prognoz=None):
    """Записать решение в момент принятия. Причина обязательна и не может быть меткой."""
    if reshenie not in ("beryom", "propuskaem", "dumaem"):
        raise ValueError("решение: beryom | propuskaem | dumaem")
    if len(prichina.strip()) < 25:
        raise ValueError("причина слишком короткая: через полгода она ничего не объяснит")
    c = _db()
    c.execute("INSERT INTO resheniya (data, obekt, reshenie, prichina, prognoz) "
              "VALUES (?,?,?,?,?)",
              (datetime.date.today().isoformat(), obekt, reshenie, prichina.strip(), prognoz))
    c.commit()
    nomer = c.execute("SELECT last_insert_rowid()").fetchone()[0]
    c.close()
    return nomer


def zakryt(obekt, itog, fakt=None):
    """Дописать фактический исход к последнему открытому решению по объекту."""
    c = _db()
    row = c.execute("SELECT id FROM resheniya WHERE obekt=? AND itog IS NULL "
                    "ORDER BY id DESC LIMIT 1", (obekt,)).fetchone()
    if not row:
        c.close()
        return False
    c.execute("UPDATE resheniya SET itog=?, fakt=?, data_itoga=? WHERE id=?",
              (itog, fakt, datetime.date.today().isoformat(), row[0]))
    c.commit()
    c.close()
    return True


def svodka():
    """Насколько прогнозы совпадали с фактом. Это и есть метрика качества системы."""
    c = _db()
    vsego = c.execute("SELECT COUNT(*) FROM resheniya").fetchone()[0]
    zakrytyh = c.execute("SELECT COUNT(*) FROM resheniya WHERE itog IS NOT NULL").fetchone()[0]
    otkrytyh = vsego - zakrytyh
    po_resheniyam = dict(c.execute(
        "SELECT reshenie, COUNT(*) FROM resheniya GROUP BY reshenie").fetchall())
    pary = c.execute("SELECT prognoz, fakt FROM resheniya "
                     "WHERE prognoz IS NOT NULL AND fakt IS NOT NULL").fetchall()
    c.close()
    oshibka = None
    if pary:
        oshibka = sum(abs(p - f) for p, f in pary) / len(pary)
    return {"всего": vsego, "закрытых": zakrytyh, "открытых": otkrytyh,
            "по решениям": po_resheniyam, "средняя ошибка прогноза": oshibka,
            "пар для сверки": len(pary)}


def main():
    s = svodka()
    print("решений всего: %(всего)d, закрыто: %(закрытых)d, ждут исхода: %(открытых)d" % s)
    for k, v in (s["по решениям"] or {}).items():
        print("   %-12s %d" % (k, v))
    if s["средняя ошибка прогноза"] is not None:
        print("средняя ошибка прогноза: %.0f (пар: %d)"
              % (s["средняя ошибка прогноза"], s["пар для сверки"]))
    else:
        print("сверять пока нечего: нет пар прогноз-факт")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
