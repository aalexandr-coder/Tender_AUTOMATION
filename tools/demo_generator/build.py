# -*- coding: utf-8 -*-
"""Генератор демонстраційної бази знань (вигадані дані).

Запуск (потрібні reportlab, python-docx, openpyxl і шрифти DejaVuSans*.ttf):

    python tools/demo_generator/build.py --out demo --fonts <папка з DejaVuSans.ttf і DejaVuSans-Bold.ttf>

Папка --out очищається перед генерацією.
"""
import argparse
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import build_company
import build_contracts
import build_knowledge
import build_tenders
import render as R
from world import COMPANY, PROJECTS, TENDERS, SPECS


def readme(out):
    R.md(out / "README.md", f"""
# Демонстраційна база знань (ДЕМО)

**Усі дані в цій папці вигадані.** Компанія «{COMPANY['short']}», замовники, особи, коди, суми, договори, акти й листи-відгуки згенеровані скриптом `tools/demo_generator/` і не мають стосунку до реальних організацій. Збіги випадкові. Кожен PDF і DOCX має позначку «ДЕМО».

Структура повторює реальну базу знань (див. `Системна інструкція проєкту (нова версія).md` і `rules/`):

| Папка | Що в ній | Кількість |
|---|---|---|
| `company/` | Паспорт, реєстр документів, витяги, статут | 6 документів + 2 реєстри |
| `finance/` | Фінзвітність 2021–2025, банківські довідки | 7 |
| `contracts/` | Договори, акти, листи-відгуки (`Оригінали/`, `Для тендеру/`), `ТЗ для ШІ.md` | {len(PROJECTS)} проєктів |
| `project_cards/` | Картки проєктів + реєстр за замовниками | {len(PROJECTS)} |
| `nuances/` | Журнал нюансів по проєктах і загальні | {len(PROJECTS)} + 3 |
| `cv/` | Базові CV, дипломи, сертифікати, архів | {len(SPECS)} спеціалістів |
| `templates/` | Шаблони CV, чек-листа, довідки про досвід | 4 |
| `tenders/` | Папки тендерів: документація, аналіз, заявка, довідки, чек-лист | {len(TENDERS)} тендерів |

Правила роботи лежать у `../rules/` (одні для реальної й демонстраційної бази).

Перегенерація: `python tools/demo_generator/build.py --out demo --fonts <шрифти DejaVu>`.
""")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="demo")
    ap.add_argument("--fonts", required=True)
    a = ap.parse_args()
    repo = Path(__file__).resolve().parents[2]
    out = (repo / a.out).resolve()
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    R.init_fonts(a.fonts)
    build_company.build(out, repo)
    build_contracts.build(out)
    build_knowledge.build_cards_and_nuances(out)
    build_knowledge.build_cv(out)
    build_tenders.build(out)
    readme(out)
    n = sum(1 for _ in out.rglob("*") if _.is_file())
    print(f"OK: {n} файлів у {out}")


if __name__ == "__main__":
    main()
