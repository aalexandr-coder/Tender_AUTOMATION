# -*- coding: utf-8 -*-
"""tenders/: п'ять вигаданих тендерів у повній структурі папки тендеру."""
import re
import shutil

import render as R
from build_contracts import folder, letter_name
from build_knowledge import cv_docx
from world import COMPANY as C, CUSTOMERS, PROJECTS, SPECS, TENDERS

TDATE = {"T1": "12.11.2026", "T2": "16.11.2026", "T3": "17.11.2026", "T4": "18.11.2026", "T5": "13.10.2026"}
BUDGET_NUM = {"T1": 9500000.0, "T2": 140000.0, "T3": 900000.0, "T4": 1800000.0, "T5": 210000.0}
RATE = {"UAH": (1600, 2600), "USD": (38, 62)}          # погодинні ставки (мін, макс)
STATUS_BY_TENDER = {
    "T1": {"adm": "В роботі", "exp": "На перевірці", "team": "В роботі", "tech": "На перевірці", "price": "Не розпочато", "dps": "В роботі"},
    "T2": {"adm": "В роботі", "exp": "На перевірці", "team": "В роботі", "tech": "В роботі", "price": "В роботі", "dps": "Не застосовується"},
    "T3": {"adm": "Не розпочато", "exp": "В роботі", "team": "Потребує уточнення", "tech": "Не розпочато", "price": "Не розпочато", "dps": "Не розпочато"},
    "T4": {"adm": "На перевірці", "exp": "Готово", "team": "На перевірці", "tech": "На перевірці", "price": "В роботі", "dps": "Готово"},
    "T5": {"adm": "Готово", "exp": "Готово", "team": "Готово", "tech": "Готово", "price": "Готово", "dps": "Не застосовується"},
}
SPEC = {s["name"]: s for s in SPECS}


def parse_spec_name(n):
    return SPEC[n]


def build(out):
    root = out / "tenders"
    for k, t in TENDERS.items():
        _tender(out, root / t["dir"], k, t)


# ---------------------------------------------------------------- helpers
def _t(k):
    return TENDERS[k]


def _is_en(t):
    return t["lang"] == "EN"


def _roles_rows(t):
    return [(r, req, name) for r, req, name in t["roles"]]


def _proj_rows(t, with_amount=False):
    rows = []
    for i, pk in enumerate(t["projects"], 1):
        p = PROJECTS[pk]
        cu = CUSTOMERS[p["cust"]]
        period = f"{p['date']} – {p['kp_date']}" if p["kp_date"] else f"{p['date']} – не завершено"
        proof = (letter_name(p).replace(".pdf", "") if p["letter_date"] else "[ПОТРІБНО: лист-відгук]")
        rows.append([str(i), f"Виконавець: {C['short']}<br/>Замовник: {cu['full']}<br/>Підтвердження: {proof}"
                     .replace("<br/>", "; "),
                     p.get("subject_ua", p["subject"]), "; ".join(p["scope"][:3]), period])
    return rows


# ---------------------------------------------------------------- main
def _tender(out, base, k, t):
    en = _is_en(t)
    cu = CUSTOMERS[t["cust"]]
    d = TDATE[k]
    st = STATUS_BY_TENDER[k]
    doc = base / "Документація"
    an = base / "Аналіз"
    ua = base / "Заявка" / "1 Українською (на перевірку керівнику)"
    adm, team, tech, price = ua / "Адміністративні документи", ua / "Команда", ua / "Технічна пропозиція", ua / "Цінова пропозиція"
    ref = base / "Довідки і підтвердні документи"

    # ---- Документація замовника
    _customer_docs(doc, k, t, cu, en)

    # ---- Аналіз
    R.md(an / "Аналіз тендеру.md", _analysis_md(k, t, cu, d))
    R.md(an / "Таблиця відповідності.md", _compliance_md(k, t, d))
    R.md(an / "Питання і ризики.md", _risks_md(k, t, d))
    R.md(an / "Склад команди і проєкти на погодження.md", _team_projects_md(k, t, d))
    _checklist(base / "Чек-лист тендеру.xlsx", k, t, cu)

    # ---- Заявка (1 Українською)
    R.docx_file(adm / f"Заявка учасника від {d}.docx", [
        ("title", "ЗАЯВКА УЧАСНИКА (форма замовника, чернетка укр.)"),
        ("p", f"Тендер: {t['subject']} ({t['ref']}). Замовник: {cu['full']}."),
        ("table", None, [["Учасник", C["full"]], ["Код ЄДРПОУ", C["edrpou"]], ["Адреса", C["address"]],
                         ["Підписант", f"Директор {C['director']}"], ["Банківський рахунок",
                                                                      "[ПОТРІБНО: вибір рахунка користувачем, rules/Уточнення під кожен тендер.md]"],
                         ["Строк дії пропозиції", "90 днів з дати подання"]]),
        ("note", "Усі твердження заяви підтверджує директор. Чернетка, потребує перевірки керівником."),
    ])
    R.docx_file(adm / f"Довідка про аналогічний досвід від {d}.docx", [
        ("title", f"ДОВІДКА-ПОРТФОЛІО проєктів {C['short']} для підтвердження досвіду"),
        ("p", f"Вимога ТД: {t['exp_req']}."),
        ("table", ["№", "Виконавець / Замовник / Підтвердження", "Предмет", "Опис", "Період"], _proj_rows(t)),
        ("note", "Без ЄДРПОУ сторін, без вартості (правило rules/Довідка про аналогічний досвід.md)."),
    ], landscape=True)
    R.docx_file(adm / f"Лист про відсутність підстав від {d}.docx", [
        ("title", "ЛИСТ-ПІДТВЕРДЖЕННЯ"),
        ("p", f"{C['short']} підтверджує: відсутня заборгованість зі сплати податків; не відкрито провадження про банкрутство; "
              "відсутні підстави для відмови в участі, визначені документацією; санкції до учасника не застосовано."),
        ("p", f"Директор {C['director']} [підпис]")])
    for role, req, name in t["roles"]:
        s = SPEC[name]
        cv_docx(team / f"CV {name.split()[0]} під тендер від {d}.docx", s, d)
    R.docx_file(tech / f"Технічна пропозиція (чернетка) від {d}.docx", _tech_blocks(k, t))
    _price_xlsx(price / f"Цінова пропозиція від {d}.xlsx", k, t)

    # ---- Заявка (2 мова тендеру) та (3 до подання)
    if en:
        en_dir = base / "Заявка" / "2 English (переклад після погодження)"
        R.docx_file(en_dir / "Адміністративні документи" / "Proposal Submission Form (EN).docx", [
            ("title", "PROPOSAL SUBMISSION FORM"),
            ("p", f"To: {cu['full']}. Re: {t['ref']} — {t['subject']}."),
            ("table", None, [["Bidder", C["en_full"]], ["Registration code", C["edrpou"]], ["Address", C["address"]],
                             ["Authorised signatory", f"Director {C['director']}"], ["Proposal validity", "90 days"]]),
            ("note", "Translation prepared after approval of the Ukrainian version.")])
        R.docx_file(en_dir / "Технічна пропозиція" / "Technical Proposal (EN).docx", _tech_blocks(k, t, en=True))
        _price_xlsx(en_dir / "Цінова пропозиція" / "Financial Offer (EN).xlsx", k, t, en=True)
    if k == "T5":
        fin = base / "Заявка" / "3 До подання (English, подано 14.10.2026)"
        for nm in ("Technical section/Annex 2 - Proposal Submission Form (signed)", "Technical section/Annex 3 - Technical Proposal (signed)",
                   "Technical section/Organizational profile", "Financial section/Annex 4 - Financial Offer (signed)"):
            R.pdf(fin / f"{nm}.pdf", [("h1", nm.split("/")[-1].upper()), ("c", f"{C['en_short']} · {t['ref']}"),
                                      ("p", "Final signed version as submitted on 14.10.2026 (demo copy: contents omitted).")],
                  title=nm)
    elif k == "T4":
        R.keep(base / "Заявка" / "3 До подання (укр., підписані)" / ".gitkeep")

    # ---- Довідки і підтвердні документи
    if st["dps"] != "Не застосовується":
        if st["dps"] in ("Готово", "В роботі"):
            R.pdf(ref / f"Довідка ДПС про відсутність заборгованості від {d}.pdf", [
                ("h1", "ДОВІДКА ПРО ВІДСУТНІСТЬ ЗАБОРГОВАНОСТІ (демо)"), ("c", f"від {d}"),
                ("p", f"{C['short']} (код {C['edrpou']}) станом на {d} податкової заборгованості не має. "
                      "Строк дії довідки — 10 календарних днів.")], title="Довідка ДПС (демо)")
    ref.mkdir(parents=True, exist_ok=True)
    shutil.copy2(out / "company" / f"Витяг ЄДР від 12.11.2026.pdf", ref / "Витяг ЄДР від 12.11.2026.pdf")
    for pk in t["projects"]:
        p = PROJECTS[pk]
        if p["letter_date"]:
            src = out / "contracts" / folder(p) / "Для тендеру" / letter_name(p)
            dst = ref / "Листи-відгуки" / f"{p['fname']} - {letter_name(p)}"
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
    for _, _, name in t["roles"]:
        s = SPEC[name]
        dst = ref / "Дипломи" / name / f"Диплом від {s['diploma_date']}.pdf"
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(out / "cv" / name / f"Диплом від {s['diploma_date']}.pdf", dst)

    R.md(base / "00 Статус підготовки.md", _status_md(k, t, cu, d, st))


# ---------------------------------------------------------------- документація замовника
def _customer_docs(doc, k, t, cu, en):
    if en:
        R.docx_file(doc / "Загальні" / f"{t['ref']} Instructions to Bidders.docx", [
            ("title", f"REQUEST FOR PROPOSAL {t['ref']}"), ("p", f"Issued by: {cu['full']}. Subject: {t['subject']}."),
            ("h1", "1. Key dates"), ("table", ["Item", "Date"], [["Clarification requests", "7 days before the deadline"],
                                                                 ["Submission deadline", f"{t['deadline']}, {t['deadline_time']} (Kyiv time)"]]),
            ("h1", "2. Submission"), ("p", f"Platform: {t['platform']}. Currency: {t['currency']}. Estimated budget: {t['budget']}."),
            ("h1", "3. Qualification"), ("bullets", [t["exp_req"], "Valid registration extract", "No tax arrears",
                                                       "Declaration of no conflict of interest"]),
            ("h1", "4. Evaluation"), ("table", ["Criterion", "Weight"], [["Technical proposal", "60%"], ["Team and experience", "20%"],
                                                                        ["Price", "20%"]]),
            ("note", "DEMO DOCUMENT: fictional tender.")])
        R.docx_file(doc / "Додатки до документації" / "Terms of Reference.docx", [
            ("title", "TERMS OF REFERENCE"), ("p", t["subject"]),
            ("bullets", ["Functional scope: see Annex A", "Accessibility: WCAG 2.1 AA", "Hosting: customer cloud",
                         "Delivery within 6 months from signing"]),
            ("h1", "Required team"), ("table", ["Role", "Minimum requirements"], [[r, q] for r, q, _ in t["roles"]])])
        R.docx_file(doc / "Додатки до документації" / "Draft Contract.docx", [
            ("title", "DRAFT CONTRACT"), ("p", "General terms: payment in milestones, IP transferred on acceptance, 6-month warranty.")])
        R.docx_file(doc / "Форми замовника (English)" / "Annex 2 - Proposal Submission Form.docx", [
            ("title", "ANNEX 2: PROPOSAL SUBMISSION FORM"), ("table", None, [["Bidder", "[ ]"], ["Address", "[ ]"],
                                                                             ["Signatory", "[ ]"]])])
        R.docx_file(doc / "Форми замовника (English)" / "Annex 3 - Past Experience Form.docx", [
            ("title", "ANNEX 3: PAST EXPERIENCE"), ("table", ["Project", "Client", "Value", "Period", "Reference"],
                                                    [["[ ]"] * 5])])
    else:
        R.docx_file(doc / "Загальні" / f"Тендерна документація {t['ref']}.docx", [
            ("title", f"ТЕНДЕРНА ДОКУМЕНТАЦІЯ {t['ref']}"),
            ("p", f"Замовник: {cu['full']}. Предмет закупівлі: {t['subject']}."),
            ("h1", "1. Строки"), ("table", ["Подія", "Дата"], [["Запити на роз'яснення", "за 5 днів до кінця строку"],
                                                               ["Кінцевий строк подання", f"{t['deadline']}, {t['deadline_time']} (за Києвом)"]]),
            ("h1", "2. Подання"), ("p", f"Спосіб: {t['platform']}. Очікувана вартість: {t['budget']}."),
            ("h1", "3. Кваліфікаційні вимоги"),
            ("bullets", [t["exp_req"], "Витяг з ЄДР (не старший за 30 днів)", "Довідка про відсутність податкового боргу",
                         "Лист про відсутність підстав для відмови в участі"]),
            ("h1", "4. Критерії оцінки"), ("table", ["Критерій", "Вага"], [["Ціна", "60%"], ["Досвід і команда", "25%"],
                                                                         ["Технічна пропозиція", "15%"]]),
            ("note", "ДЕМО-ДОКУМЕНТ: тендер вигаданий.")])
        R.docx_file(doc / "Додатки до документації" / "Технічне завдання замовника.docx", [
            ("title", "ТЕХНІЧНЕ ЗАВДАННЯ"), ("p", t["subject"]),
            ("bullets", ["Обсяг функціоналу — Додаток А", "Вебдоступність WCAG 2.1 AA", "Хостинг — інфраструктура замовника",
                         "Строк виконання — до 6 місяців з дати підписання"]),
            ("h1", "Склад команди"), ("table", ["Роль", "Мінімальні вимоги"], [[r, q] for r, q, _ in t["roles"]])])
        R.docx_file(doc / "Додатки до документації" / "Проєкт договору.docx", [
            ("title", "ПРОЄКТ ДОГОВОРУ"), ("p", "Оплата за етапами, права на програму переходять після приймання, гарантія 6 місяців.")])
        R.docx_file(doc / "Форми замовника (українська)" / "Форма 1 - Заявка учасника.docx", [
            ("title", "ФОРМА 1: ЗАЯВКА УЧАСНИКА"), ("table", None, [["Учасник", "[ ]"], ["Адреса", "[ ]"], ["Підписант", "[ ]"]])])
        R.docx_file(doc / "Форми замовника (українська)" / "Форма 2 - Довідка про досвід.docx", [
            ("title", "ФОРМА 2: ДОВІДКА ПРО ДОСВІД"), ("table", ["Проєкт", "Замовник", "Період", "Підтвердження"], [["[ ]"] * 4])])


# ---------------------------------------------------------------- аналіз
def _analysis_md(k, t, cu, d):
    reqs = "\n".join(f"- **{r}** — {q}" for r, q, _ in t["roles"])
    return f"""
# Аналіз тендеру: {t['subject']}

> ДЕМО-ДОКУМЕНТ. Тендер вигаданий. Складено {d} за `rules/Папка тендеру.md`.

| Поле | Значення |
|---|---|
| Замовник | {cu['full']} ({cu['kind']}) |
| Референс | {t['ref']} |
| Предмет | {t['subject']} |
| Майданчик / спосіб подання | {t['platform']} |
| Дедлайн | {t['deadline']}, {t['deadline_time']} (за Києвом) |
| Орієнтовна вартість | {t['budget']} |
| Валюта | {t['currency']} |
| Форма пропозицій | {t['envelopes']} |
| Мова заявки | {'англійська (спочатку готуємо українською)' if _is_en(t) else 'українська'} |

## Кваліфікаційні вимоги

- Досвід: {t['exp_req']}.
- Витяг з ЄДР; довідка про відсутність податкового боргу; лист про відсутність підстав для відмови.

## Команда

{reqs}

## Критерії оцінки

Ціна, досвід і команда, технічна пропозиція (ваги — у тендерній документації).

## Що готуємо

Заявка учасника, довідка про досвід (довідка-портфоліо), CV команди під тендер, технічна пропозиція, цінова пропозиція.
"""


def _compliance_md(k, t, d):
    st = STATUS_BY_TENDER[k]
    rows = [("Витяг з ЄДР", "company/Витяг ЄДР від 12.11.2026.pdf", "є"),
            ("Відсутність податкового боргу", "Довідка ДПС під тендер", {"Готово": "є", "В роботі": "частково"}.get(st["dps"], "немає")),
            ("Досвід: " + t["exp_req"], "project_cards/ + листи-відгуки", "є" if st["exp"] in ("Готово", "На перевірці") else "частково")]
    for r, q, n in t["roles"]:
        s = SPEC[n]
        ok = "є" if not (n == "Макетна Ірина" and _is_en(t)) else "частково (мовна вимога: EN B1 < B2)"
        rows.append((f"Роль «{r}»: {q}", f"cv/{n}/ (CV, диплом)", ok))
    body = "\n".join(f"| {a} | {b} | {c} |" for a, b, c in rows)
    return f"""
# Таблиця відповідності

> ДЕМО-ДОКУМЕНТ. Станом на {d}.

| Вимога | Чим підтверджуємо | Статус |
|---|---|---|
{body}
"""


def _risks_md(k, t, d):
    common = ["Підписи й заяви про відсутність підстав для відмови підтверджує директор (не ШІ)."]
    spec = {
        "T1": ["Проєкт «Платформа: Відкриті дані» не завершено (немає акта КП), тому в довідку не включено.",
               "Довідка ДПС має короткий строк дії: замовляти за 3–5 днів до подання."],
        "T2": ["Вимога EN B2 для QA; у Макетної Ірини B1: рішення за керівником (замінити або позначити ризик).",
               "Два конверти: цінову частину не додавати в технічний конверт."],
        "T3": ["Немає спеціаліста з сертифікованим пентестом; сертифікат Kubernetes Богдана Тестового прострочено (нюанс у `nuances/Спеціалісти.md`).",
               "Дедлайн 10.12.2026; є час на запит роз'яснень."],
        "T4": [f"Дедлайн {t['deadline']}: залишилось 9 днів. Пріоритет: довідка ДПС і підпис КЕП директора.",
               "Лист-відгук за «Реєстр: Ліцензії» без КЕП (нюанс): уточнити, чи достатньо скану."],
        "T5": ["Заявку подано 14.10.2026; папка лишається зразком (приклад) для наступних тендерів.",
               "Курсові різниці: ціна в USD, договір «Сервіс-Ветеран» у гривні."],
    }[k]
    items = "\n".join(f"{i}. {x}" for i, x in enumerate(common + spec, 1))
    return f"# Питання і ризики\n\n> ДЕМО-ДОКУМЕНТ. Оновлено {d}.\n\n{items}\n"


def _team_projects_md(k, t, d):
    team = "\n".join(f"| {r} | {n} | {SPEC[n]['role']} |" for r, _, n in t["roles"])
    pr = "\n".join(f"| {PROJECTS[p]['internal']} | {PROJECTS[p]['date']} | {PROJECTS[p]['kp_date'] or 'не завершено'} |"
                   for p in t["projects"])
    return f"""
# Склад команди і проєкти на погодження

> ДЕМО-ДОКУМЕНТ. Станом на {d}. Погодження: [ПОТРІБНО: Head of Legal].

## Команда

| Роль за ТД | Спеціаліст | Посада в компанії |
|---|---|---|
{team}

## Проєкти для підтвердження досвіду

| Проєкт | Дата договору | Акт КП |
|---|---|---|
{pr}
"""


def _status_md(k, t, cu, d, st):
    stages = [("1. Аналіз тендеру", "зроблено"), ("2. Звірка чек-листа з ТД", "зроблено"),
              ("3. Таблиця відповідності", "зроблено" if st["exp"] in ("Готово", "На перевірці") else "в роботі"),
              ("4. Погодження проєктів і команди", {"Готово": "зроблено", "На перевірці": "на перевірці"}.get(st["team"], "в роботі")),
              ("5. Адміністративні документи, CV", st["adm"].lower()), ("6. Технічна пропозиція", st["tech"].lower()),
              ("7. Цінова пропозиція", st["price"].lower()),
              ("8. Переклад (якщо потрібен)", "після погодження" if _is_en(t) and k != "T5" else ("зроблено" if k == "T5" else "не потрібен")),
              ("9. Підписання і подання", "подано 14.10.2026" if k == "T5" else "—")]
    sr = "\n".join(f"| {a} | {b} |" for a, b in stages)
    return f"""
# Статус підготовки: {t['subject']} ({t['ref']})

> ДЕМО-ДОКУМЕНТ. Оновлено {d}. Веде ШІ (`rules/Папка тендеру.md`).

**Дедлайн подання:** {t['deadline']}, {t['deadline_time']}. **Статус:** {t['status']}.
**Замовник:** {cu['full']}. **Подання:** {t['platform']}; {t['envelopes']}.
**Мова заявки:** {'англійська' if _is_en(t) else 'українська'}.

## Карта папки

| Папка | Що там |
|---|---|
| `Документація/` | Файли замовника без змін (Загальні, Додатки, Форми) |
| `Аналіз/` | Аналіз тендеру, таблиця відповідності, питання і ризики, склад команди і проєкти на погодження |
| `Чек-лист тендеру.xlsx` | Робочий чек-лист (аркуші Чек-лист, Досвід, Команда) |
| `Заявка/1 Українською (на перевірку керівнику)/` | Адміністративні документи, Команда, Технічна пропозиція, Цінова пропозиція |
| `Заявка/2 …` | {'Переклад англійською після погодження' if _is_en(t) else 'Не потрібна: тендер українською'} |
| `Заявка/3 До подання …` | {'Фінальні підписані файли' if k in ('T5',) else 'Поки порожня'} |
| `Довідки і підтвердні документи/` | Витяг ЄДР, довідка ДПС (за потреби), листи-відгуки, дипломи |

## Етапи

| Етап | Статус |
|---|---|
{sr}
"""


# ---------------------------------------------------------------- технічна пропозиція
def _tech_blocks(k, t, en=False):
    p = PROJECTS[t["projects"][0]]
    plan = [["1", "Аналіз і проєктування", "4–6 тижнів"], ["2", "Розроблення", "10–14 тижнів"],
            ["3", "Тестування, безпека, доступність", "4 тижні"], ["4", "Впровадження і навчання", "2–3 тижні"]]
    if en:
        return [("title", "TECHNICAL PROPOSAL"), ("p", f"{t['subject']} · {t['ref']} · {C['en_short']}"),
                ("h1", "1. Understanding of the assignment"), ("p", "We understand the objective as delivering a secure, accessible, maintainable solution."),
                ("h1", "2. Approach"), ("p", f"Iterative delivery (Scrum), CI/CD, security by design. Reference stack: {p['tech']}."),
                ("h1", "3. Work plan"), ("table", ["#", "Phase", "Duration"], plan),
                ("h1", "4. Team"), ("table", ["Role", "Specialist"], [[r, n] for r, _, n in t["roles"]]),
                ("note", "Translation prepared after approval of the Ukrainian version; technical content validated by PM.")]
    return [("title", "ТЕХНІЧНА ПРОПОЗИЦІЯ (чернетка)"), ("p", f"{t['subject']} · {t['ref']} · {C['short']}"),
            ("h1", "1. Розуміння завдання"), ("p", "Метою є створення безпечного, доступного й підтримуваного рішення відповідно до ТЗ замовника."),
            ("h1", "2. Підхід і методологія"), ("p", f"Ітеративна розробка (Scrum), CI/CD, безпека за замовчуванням. Опорний стек з аналогічного проєкту «{p['internal']}»: {p['tech']}."),
            ("h1", "3. План робіт"), ("table", ["№", "Етап", "Тривалість"], plan),
            ("h1", "4. Команда"), ("table", ["Роль", "Спеціаліст"], [[r, n] for r, _, n in t["roles"]]),
            ("h1", "5. Ризики"), ("bullets", ["Зміна вимог — управління через backlog і погодження змін",
                                               "Інтеграції зі сторонніми реєстрами — ранні тести шлюзів"]),
            ("note", "Валідація PM до [дата]. ШІ технічних рішень не вигадує; зміст дає PM.")]


# ---------------------------------------------------------------- ціна
def _price_xlsx(path, k, t, en=False):
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Calculation" if en else "Калькуляція"
    hdr = ["Role", "Hours", "Rate", "Amount"] if en else ["Роль", "Години", "Ставка за годину", "Сума"]
    ws.append([("FINANCIAL OFFER" if en else "ЦІНОВА ПРОПОЗИЦІЯ") + f" — {t['ref']} (DEMO)"])
    ws.append([f"{C['en_short'] if en else C['short']} · {t['currency']}"])
    ws.append([])
    ws.append(hdr)
    for c in ws[4]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="E6EBF2")
    lo, hi = RATE[t["currency"]]
    n = len(t["roles"])
    target = BUDGET_NUM[k] * 0.92
    base_rates = [hi - i * (hi - lo) // max(1, n) for i in range(n)]
    weights = [1.0 + 0.15 * i for i in range(n)]
    total_w = sum(w * r for w, r in zip(weights, base_rates))
    scale = target / total_w
    for i, (r, _, _) in enumerate(t["roles"]):
        hours = max(40, round(weights[i] * scale / 10) * 10)
        row = 5 + i
        ws.append([r, hours, base_rates[i], f"=B{row}*C{row}"])
    last = 4 + n
    ws.append(["Total" if en else "Разом", f"=SUM(B5:B{last})", "", f"=SUM(D5:D{last})"])
    ws.append(["VAT" if en else "ПДВ", "", "", "included" if (en and t["currency"] != "UAH") else ("з ПДВ" if t["currency"] == "UAH" else "без ПДВ")])
    ws.append(["Estimated budget" if en else "Орієнтовна вартість за ТД", "", "", t["budget"]])
    ws.column_dimensions["A"].width = 38
    for col in "BCD":
        ws.column_dimensions[col].width = 20
    ws[f"A{last + 1}"].font = Font(bold=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))


# ---------------------------------------------------------------- чек-лист
def _checklist(path, k, t, cu):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.worksheet.datavalidation import DataValidation
    st = STATUS_BY_TENDER[k]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Чек-лист"
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="E6EBF2")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    ws["A1"] = "ЧЕК-ЛИСТ ТЕНДЕРУ (ДЕМО)"
    ws["A1"].font = Font(bold=True, size=14)
    gen = [("Назва тендеру", f"{t['subject']} ({t['ref']})"), ("Замовник / платник", cu["full"]),
           ("Майданчик / спосіб подання", t["platform"]), ("Офіційний строк подання", f"{t['deadline']}, {t['deadline_time']} (Київ)"),
           ("Валюта / ПДВ", f"{t['currency']}; {t['budget']}"), ("Форма пропозицій", t["envelopes"]),
           ("Підписання", "КЕП директора"), ("Відповідальний юрист", "Юрист")]
    for i, (a, b) in enumerate(gen):
        ws.cell(5 + i, 1, a).font = bold
        ws.cell(5 + i, 3, b).fill = yellow
    ws["G5"], ws["G6"], ws["G7"], ws["G8"], ws["G9"] = "Усього вимог", "Готово", "Не застосовується", "Потребує уточнення", "Готовність, %"
    ws["H5"], ws["H6"] = "=COUNT(A17:A67)", '=COUNTIF(F17:F67,"Готово")'
    ws["H7"], ws["H8"] = '=COUNTIF(F17:F67,"Не застосовується")', '=COUNTIF(F17:F67,"Потребує уточнення")'
    ws["H9"] = "=IFERROR(H6/(H5-H7),0)"
    ws["H9"].number_format = "0%"
    hdr = ["№", "Пункт ТД", "Вимога", "Яким документом покривається", "Наш документ", "Статус", "Відповідальний", "Примітка"]
    for i, h in enumerate(hdr, 1):
        c = ws.cell(15, i, h)
        c.font = bold
        c.fill = fill
    rows = [("1. Кваліфікаційні вимоги", None),
            ("п. 3.1", "Учасник зареєстрований як юридична особа", "Витяг з ЄДР", "Витяг ЄДР від 12.11.2026.pdf", "Готово", "Юрист", ""),
            ("п. 3.2", "Відсутність податкового боргу", "Довідка ДПС", f"Довідка ДПС від {TDATE[k]}.pdf" if st["dps"] in ("Готово", "В роботі") else "", st["dps"], "Юрист", "Короткий строк дії"),
            ("п. 3.3", "Відсутність підстав для відмови в участі", "Лист у довільній формі", "Лист про відсутність підстав.docx", st["adm"], "Юрист", "Підтверджує директор"),
            ("2. Досвід Компанії", None),
            ("п. 4.1", t["exp_req"], "Довідка про досвід + листи-відгуки", "Довідка про аналогічний досвід.docx", st["exp"], "Юрист", "Перелік — аркуш «Досвід»"),
            ("3. Склад команди", None)]
    for i, (r, q, n) in enumerate(t["roles"]):
        rows.append((f"ТЗ, склад команди", f"{r}: {q}", "Резюме, диплом", f"CV {n.split()[0]} під тендер.docx", st["team"], "Юрист", n))
    rows += [("4. Технічна пропозиція", None),
             ("ТЗ", "Технічна пропозиція за вимогами ТЗ", "Технічна пропозиція", "Технічна пропозиція.docx", st["tech"], "Юрист / PM", "Валідація PM до [дата]"),
             ("5. Цінова пропозиція", None),
             ("Форма ціни", "Цінова пропозиція за формою замовника", "Калькуляція", "Цінова пропозиція.xlsx", st["price"], "Юрист", "Бюджет — від DM")]
    r = 16
    num = 0
    sec_rows = []
    ws.cell(r, 1, "Наскрізна нумерація").font = Font(italic=True)
    r = 17
    for item in rows:
        if item[1] is None:
            ws.cell(r, 1, item[0]).font = bold
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        else:
            num += 1
            ws.cell(r, 1, float(f"{num}.0"))
            for j, v in enumerate(item, 2):
                ws.cell(r, j, v)
            sec_rows.append(r)
        r += 1
    for col, w in zip("ABCDEFGH", (8, 16, 52, 34, 40, 20, 16, 34)):
        ws.column_dimensions[col].width = w
    wsd = wb.create_sheet("Довідники")
    for i, v in enumerate(["Статуси", "Не розпочато", "В роботі", "На перевірці", "Готово", "Потребує уточнення", "Не застосовується"], 1):
        wsd.cell(i, 1, v)
    dv = DataValidation(type="list", formula1="'Довідники'!$A$2:$A$7", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("F17:F67")

    we = wb.create_sheet("Досвід", 1)
    we["A1"] = "ДОСВІД КОМПАНІЇ"
    we["A1"].font = Font(bold=True, size=14)
    we["A4"], we["C4"] = "Вимоги ТД до досвіду", t["exp_req"]
    we["A6"], we["C6"] = "Перелік погоджено Head of Legal", "[ПОТРІБНО: дата погодження]"
    for i, h in enumerate(["№", "Назва проєкту", "Підтвердження", "Посилання на підтвердження", "Вимога ТД, яку покриває", "Статус", "Примітка"], 1):
        c = we.cell(8, i, h)
        c.font = bold
        c.fill = fill
    for i, pk in enumerate(t["projects"], 1):
        p = PROJECTS[pk]
        proof = "Лист-відгук" if p["letter_date"] else "—"
        we.append([f"{i}.0", f"Договір № {p['contract_no']} від {p['date']}: {p['internal']}", proof,
                   letter_name(p) if p["letter_date"] else "", "Досвід у розробленні аналогічних систем",
                   "Готово" if p["letter_date"] else "Потребує уточнення", "Суми в актах закрито" if p["letter_date"] else "Немає листа-відгуку"])
    wt = wb.create_sheet("Команда", 2)
    wt["A1"] = "СКЛАД КОМАНДИ"
    wt["A1"].font = Font(bold=True, size=14)
    for i, h in enumerate(["№", "Роль за ТД", "Вимоги ТД до ролі", "ПІБ спеціаліста", "Посилання на документи", "Статус", "Примітка"], 1):
        c = wt.cell(8, i, h)
        c.font = bold
        c.fill = fill
    for i, (rname, q, n) in enumerate(t["roles"], 1):
        s = SPEC[n]
        note = ""
        if n == "Макетна Ірина" and _is_en(t):
            note = "Рівень EN B1 нижче вимоги B2: ризик"
        wt.append([f"{i}.0", rname, q, n, f"cv/{n}/CV від {s['cv_date']}.docx", st["team"], note])
    for ws_ in (we, wt):
        for col in "ABCDEFG":
            ws_.column_dimensions[col].width = 28
    wi = wb.create_sheet("Інструкція")
    for i, (a, b) in enumerate([("ПОРЯДОК ЗАПОВНЕННЯ ЧЕК-ЛИСТА", ""), ("Вимога", "Копіюється з ТД дослівно"),
                                ("Статус", "Лише зі списку аркуша «Довідники»"), ("Відповідальний", "Посада виконавця")], 1):
        wi.cell(i, 1, a).font = bold
        wi.cell(i, 2, b)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))
