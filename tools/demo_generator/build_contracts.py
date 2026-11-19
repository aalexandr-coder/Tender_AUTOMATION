# -*- coding: utf-8 -*-
"""contracts/: договори, акти, листи-відгуки, ТЗ для ШІ (усе вигадане)."""
import render as R
from world import COMPANY as C, CUSTOMERS, PROJECTS


def folder(p):
    return f"{p['fname']} від {p['date']}"


def letter_name(p):
    return f"Лист-відгук від {p['letter_date']}" + (" (КЕП)" if p["letter_kep"] else "") + ".pdf"


def contract_pdf(path, p):
    cu = CUSTOMERS[p["cust"]]
    cur = p["cur"]
    price_rows = [[n, R.money(v, cur), d or "—"] for n, v, d in p["stages"]]
    price_rows.append(["<b>Разом</b>", f"<b>{R.money(p['total'], cur)}</b> ({p['vat']})", ""])
    is_frame = p.get("frame")
    subject = p.get("subject_ua", p["subject"])
    blocks = [
        ("h1", ("ДОГОВІР" if not is_frame else "ЗАВДАННЯ ДО ДОВГОСТРОКОВОЇ УГОДИ") + f" № {p['contract_no']}"),
        ("c", f"м. Київ · {p['date']}"),
        ("p", f"{cu['full']} (далі — Замовник), в особі {cu['signer']}, з одного боку, та {C['full']} (далі — Виконавець), "
              f"в особі директора {C['director']}, який діє на підставі Статуту, з іншого боку, уклали цей договір про таке."),
        ("h2", "1. Предмет договору"),
        ("p", f"1.1. Виконавець зобов'язується надати послуги: {subject}, а Замовник — прийняти й оплатити їх."),
        ("p", "1.2. Обсяг, вимоги й результати робіт визначені в Технічному завданні (Додаток 1), що є невід'ємною частиною договору."),
        ("h2", "2. Вартість і порядок розрахунків"),
        ("p", f"2.1. Загальна вартість послуг — {R.money(p['total'], cur)} ({p['vat']})."),
        ("table", ["Етап", "Вартість", "Строк акта"], price_rows, [100, 45, 25]),
        ("p", "2.2. Оплата здійснюється протягом 20 банківських днів після підписання акта приймання відповідного етапу."),
        ("h2", "3. Строки"),
        ("p", f"3.1. Договір набуває чинності з дати підписання {p['date']} і діє до повного виконання зобов'язань."),
        ("p", "3.2. Строки етапів — відповідно до таблиці пункту 2.1."),
        ("h2", "4. Приймання результатів"),
        ("p", "4.1. Результати кожного етапу передаються за актом. Замовник у строк 10 робочих днів підписує акт або надсилає "
              "обґрунтовану відмову."),
        ("p", "4.2. Остаточна передача комп'ютерної програми, вихідних кодів і технічної документації оформлюється актом "
              "приймання-передачі комп'ютерної програми (акт КП)."),
        ("h2", "5. Права інтелектуальної власності, конфіденційність"),
        ("p", "5.1. Майнові права на створену комп'ютерну програму переходять до Замовника з моменту підписання акта КП."),
        ("p", "5.2. Сторони не розголошують конфіденційну інформацію. Виконавець має право зазначати факт виконання договору "
              "в портфоліо без розкриття вартості та технічних деталей, що становлять комерційну таємницю."),
        ("h2", "6. Гарантія"),
        ("p", f"6.1. Гарантійний строк: {p['warranty']}."),
        ("h2", "7. Відповідальність і форс-мажор"),
        ("p", "7.1. За порушення строків сторона сплачує пеню 0,1% від вартості неналежно виконаного етапу за кожний день прострочення."),
        ("p", "7.2. Сторони звільняються від відповідальності за обставин непереборної сили, підтверджених у встановленому порядку."),
        ("h2", "8. Реквізити сторін"),
        ("table", ["Замовник", "Виконавець"], [[f"{cu['full']}<br/>Код: {cu['edrpou']}<br/>{cu['address']}",
                                               f"{C['short']}<br/>Код ЄДРПОУ: {C['edrpou']}<br/>{C['address']}"]], [85, 85]),
        ("p", "Підписи сторін: [у копії без підписів]"),
    ]
    R.pdf(path, blocks, title=f"Договір {p['contract_no']} (демо)")


def act_pdf(path, p, title, dt, amount, stage_name, hide, kp=False):
    cu = CUSTOMERS[p["cust"]]
    cur = p["cur"]
    body = [
        ("h1", title), ("c", f"від {dt} · договір № {p['contract_no']} від {p['date']}"),
        ("p", f"Ми, що нижче підписалися, {cu['full']} (Замовник) та {C['full']} (Виконавець), склали цей акт про те, що "
              f"Виконавець передав, а Замовник прийняв: {stage_name}."),
    ]
    if kp:
        body.append(("p", "Передано комп'ютерну програму, вихідні коди, технічну й експлуатаційну документацію. "
                          "Права на програму переходять до Замовника відповідно до розділу 5 договору."))
    body += [
        ("p", f"Вартість за цим актом: {R.redact(R.money(amount, cur), hide)} ({p['vat']})."),
        ("p", "Сторони претензій одна до одної не мають."),
        ("table", ["Від Замовника", "Від Виконавця"], [[cu["signer"] + "<br/>[підпис]", C["director"] + "<br/>[підпис]"]], [85, 85]),
    ]
    R.pdf(path, body, title=title + " (демо)")


def letter_pdf(path, p):
    cu = CUSTOMERS[p["cust"]]
    donor = CUSTOMERS.get(p.get("donor", ""), None)
    en = p["cust"] == "DDP"
    if en:
        blocks = [("h1", "REFERENCE LETTER"), ("c", f"{p['letter_date']}"),
                  ("p", f"{cu['full']} confirms that {C['en_short']} duly performed the services under {p['contract_no']}: "
                        f"{p['subject']}. The work was completed on {p['end']} in full and within the agreed time. "
                        "The Customer has no complaints about quality of the delivered software."),
                  ("p", f"{cu['signer']}<br/>[signature]")]
    else:
        extra = f" за рахунок коштів {donor['full']}" if donor else ""
        blocks = [("h1", "ЛИСТ-ВІДГУК"), ("c", f"від {p['letter_date']}"),
                  ("p", f"{cu['full']} підтверджує, що {C['short']} належним чином виконало договір № {p['contract_no']} від "
                        f"{p['date']}: {p['subject']}{extra}. Роботи завершено {p['end']} у повному обсязі й у погоджені строки. "
                        "Претензій до якості програмного забезпечення Замовник не має."),
                  ("p", "Рекомендуємо Виконавця для участі в аналогічних закупівлях."),
                  ("p", f"{cu['signer']}<br/>[підпис" + (" КЕП" if p["letter_kep"] else "") + "]")]
    R.pdf(path, blocks, title="Лист-відгук (демо)")


def tor_md(p):
    cu = CUSTOMERS[p["cust"]]
    en = p["cust"] == "DDP"
    scope = "\n".join(f"{i}. {s}" for i, s in enumerate(p["scope"], 1))
    stages = "\n".join(f"- {n}: {R.money(v, p['cur'])}" + (f", строк {d}" if d else "") for n, v, d in p["stages"])
    return f"""
# Технічне завдання: {p['internal']}

Текстова версія для ШІ. Джерело: `Оригінали/Договір від {p['date']} (без підписів).pdf`, Додаток 1. **ДЕМО-ДОКУМЕНТ, дані вигадані.**

**Що звірено.** Номер договору {p['contract_no']} від {p['date']}; замовник — {cu['full']}; вартість {R.money(p['total'], p['cur'])} ({p['vat']}); гарантія: {p['warranty']}.

---

## 1. Мета

{'The objective is to develop and deploy: ' + p['subject'] + '.' if en else 'Метою проєкту є ' + p['subject'] + '.'}

## 2. Обсяг робіт

{scope}

## 3. Технології

{p['tech']}

## 4. Нефункціональні вимоги

- Доступність: WCAG 2.1 рівня AA, перевірка автоматизованими інструментами.
- Безпека: захищені канали, журналювання подій, розмежування прав доступу, резервне копіювання.
- Продуктивність: до 500 одночасних користувачів, час відповіді сторінки до 2 с.

## 5. Етапи і вартість

{stages}

## 6. Результати робіт

Комп'ютерна програма, вихідні коди, технічна й експлуатаційна документація, інструкції користувача, звіт про тестування.
""" if not en else f"""
# Terms of Reference: {p['internal']}

Text version for AI. Source: `Оригінали/Договір від {p['date']} (без підписів).pdf`, Annex 1. **DEMO DOCUMENT, fictional data.**

**Verified.** Contract {p['contract_no']} dated {p['date']}; customer — {cu['full']}; value {R.money(p['total'], p['cur'])} ({p['vat']}); warranty: {p['warranty']}.

---

## 1. Objective

To develop and deploy: {p['subject']}.

## 2. Scope of work

{scope}

## 3. Technologies

{p['tech']}

## 4. Non-functional requirements

- Accessibility: WCAG 2.1 level AA, automated checks.
- Security: protected channels, event logging, role-based access, backups.
- Performance: up to 500 concurrent users, page response time under 2 s.

## 5. Deliverables and price

{stages}

## 6. Results

Software, source code, technical and user documentation, user guides, test report.
"""


def build(out):
    root = out / "contracts"
    R.keep(root / "_на_видалення" / ".gitkeep")
    R.keep(root / "_очікують_договору" / ".gitkeep")
    for p in PROJECTS.values():
        base = root / folder(p)
        orig, forten = base / "Оригінали", base / "Для тендеру"
        contract_pdf(orig / f"Договір від {p['date']} (без підписів).pdf", p)
        if p["key"] == "P1":
            R.pdf(orig / "Додаткова угода № 1 від 15.05.2024.pdf", [
                ("h1", "ДОДАТКОВА УГОДА № 1"), ("c", f"до договору № {p['contract_no']} від {p['date']} · 15.05.2024"),
                ("p", "Сторони погодили перенесення строку етапу 1 на 28.06.2024 без зміни вартості договору."),
            ], title="Додаткова угода (демо)")
        if p.get("frame"):
            R.pdf(orig / "Рамковий договір від 14.06.2022.pdf", [
                ("h1", "LONG-TERM AGREEMENT DDP-LTA-2022-045"), ("c", "14.06.2022"),
                ("p", "Framework terms for software development services. Rates, general terms and conditions: see Annexes 1–3. "
                      "(Demo copy, fictional data.)")], title="LTA (демо)")
        for n, (sname, price, dt) in enumerate(p["stages"], 1):
            if not dt:
                continue
            is_final = (dt == p["kp_date"])
            if is_final:
                continue
            title = "АКТ ЗДАЧІ-ПРИЙМАННЯ" + (f" № {n}" if p.get("frame") else "") if p.get("frame") else "АКТ НАДАННЯ ПОСЛУГ"
            fn = (f"Акт здачі-приймання № {n} від {dt}.pdf" if p.get("frame") else f"Акт надання послуг від {dt} за етапом {n}.pdf")
            fn_red = fn.replace(".pdf", " (замазано).pdf")
            act_pdf(orig / fn, p, title, dt, price, sname, False)
            act_pdf(forten / fn_red, p, title, dt, price, sname, True)
        if p["kp_date"]:
            last = p["stages"][-1]
            kp_name = (f"Акт здачі-приймання № {len(p['stages'])} від {p['kp_date']}.pdf" if p.get("frame")
                       else f"Акт приймання-передачі КП від {p['kp_date']}.pdf")
            act_pdf(orig / kp_name, p, "АКТ ПРИЙМАННЯ-ПЕРЕДАЧІ КОМП'ЮТЕРНОЇ ПРОГРАМИ", p["kp_date"], p["total"], last[0], False, kp=True)
            act_pdf(forten / kp_name.replace(".pdf", " (замазано).pdf"), p, "АКТ ПРИЙМАННЯ-ПЕРЕДАЧІ КОМП'ЮТЕРНОЇ ПРОГРАМИ",
                    p["kp_date"], p["total"], last[0], True, kp=True)
        if p["letter_date"]:
            letter_pdf(orig / letter_name(p), p)
            letter_pdf(forten / letter_name(p), p)
        R.md(base / "ТЗ для ШІ.md", tor_md(p))
