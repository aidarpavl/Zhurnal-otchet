import io
import re
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import pdfplumber
from openpyxl.styles import Font, PatternFill, Alignment

# ============================================================
# НАСТРОЙКА
# ============================================================
st.set_page_config(page_title="Журнал — Алгебра 10 А", page_icon="📘", layout="wide")

TOPICS = [
    "Повторение курса 9 класса",
    "Функции и графики",
    "Тригонометрические функции",
    "Тригонометрические уравнения",
    "Преобразование выражений",
    "Производная",
    "Применение производной",
    "Итоговое повторение",
]

QUESTION_BANK = {
    "Повторение курса 9 класса": [
        "Решите уравнение x² − 5x + 6 = 0.",
        "Упростите выражение (a+b)² − (a−b)².",
        "Найдите область определения функции y = √(x−3).",
        "Решите неравенство 2x − 7 > 0.",
        "Постройте график y = |x| и опишите его свойства.",
    ],
    "Функции и графики": [
        "Найдите область значений функции y = x² − 4x + 5.",
        "Определите чётность функции y = x³ − x.",
        "Постройте график y = 1/x и опишите его асимптоты.",
        "Найдите нули функции y = x² − 9.",
        "Как из графика y = f(x) получить y = f(x−2)+3?",
    ],
    "Тригонометрические функции": [
        "Вычислите sin(π/6) + cos(π/3).",
        "Определите период функции y = sin(2x).",
        "Найдите область значений y = 3cos x − 1.",
        "Постройте график y = tg x на [0; 2π].",
        "Выразите cos²x через sin²x.",
    ],
    "Тригонометрические уравнения": [
        "Решите уравнение sin x = 1/2.",
        "Решите уравнение cos x = 0.",
        "Решите уравнение tg x = √3.",
        "Решите 2sin²x − sin x − 1 = 0.",
        "Решите sin x + cos x = 0.",
    ],
    "Преобразование выражений": [
        "Упростите (sin x + cos x)².",
        "Докажите тождество 1 + tg²x = 1/cos²x.",
        "Упростите sin(π − x).",
        "Разложите на множители sin²x − cos²x.",
        "Вычислите cos 2x, если sin x = 0,6.",
    ],
    "Производная": [
        "Найдите производную f(x)=x³−2x.",
        "Найдите производную f(x)=sin x · x.",
        "Найдите производную f(x)=(2x+1)⁵.",
        "Найдите производную f(x)=ln x + eˣ.",
        "Найдите f′(1), если f(x)=x²+3x.",
    ],
    "Применение производной": [
        "Найдите точки экстремума f(x)=x³−3x.",
        "Найдите промежутки возрастания f(x)=x²−4x.",
        "Напишите уравнение касательной к y=x² в точке x=1.",
        "Найдите наибольшее значение f(x)=−x²+4x на [0;3].",
        "Исследуйте функцию y=x³−3x² и постройте график.",
    ],
    "Итоговое повторение": [
        "Решите систему: { x+y=5; x−y=1 }.",
        "Упростите (a−b)(a+b) − a².",
        "Решите уравнение 2ˣ = 8.",
        "Найдите производную f(x)=cos x.",
        "Решите неравенство x² − 4 < 0.",
    ],
}

# ============================================================
# ПАРСИНГ PDF — БЕЗ OCR
# ============================================================
NAME_RX = re.compile(
    r"([А-ЯЁӘҒҚҢӨҰҮҺІ][а-яёәғқңөұүһі\-]+(?:\s+[А-ЯЁӘҒҚҢӨҰҮҺІ][а-яёәғқңөұүһі\-]+){1,3})"
)
MARK_RX = re.compile(r"^(1|2|3|4|5|10|А|С|Н)$", re.IGNORECASE)


def _extract_from_lines(text: str) -> list:
    """Парсит строки вида 'ФИО 4 5 3 4 А 5 4 4'."""
    out = []
    seen = set()
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if re.search(r"учител|мұғалім|teacher|тоқсан|четверть|quarter", line, re.I):
            continue
        if re.fullmatch(r"[\d\s\.\-\/№#]+", line):
            continue
        m = NAME_RX.search(line)
        if not m:
            continue
        name = m.group(1).strip()
        if len(name.replace(" ", "")) < 4:
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)

        tail = line[m.end():]
        tokens = re.split(r"[\s,;|\t]+", tail)
        marks = [t.upper() for t in tokens if MARK_RX.match(t)]
        out.append({"name": name, "marks": marks})
    return out


def _extract_from_tables(tables: list) -> list:
    """Парсит таблицу, извлечённую pdfplumber.extract_tables()."""
    out = []
    seen = set()
    for table in tables:
        for row in table:
            if not row:
                continue
            # Убираем None и лишние пробелы
            cells = [(c or "").strip() for c in row]
            # Ищем ФИО в первых двух ячейках
            name = ""
            name_idx = -1
            for i, c in enumerate(cells[:3]):
                m = NAME_RX.search(c)
                if m:
                    name = m.group(1).strip()
                    name_idx = i
                    break
            if not name or len(name.replace(" ", "")) < 4:
                continue
            if re.search(r"учител|мұғалім|teacher", name, re.I):
                continue
            key = name.lower()
            if key in seen:
                continue
            seen.add(key)

            # Оценки — из ячеек после имени
            marks = []
            for c in cells[name_idx + 1:]:
                if not c:
                    continue
                # В одной ячейке может быть несколько токенов
                for tok in re.split(r"[\s,;|\t]+", c):
                    if MARK_RX.match(tok):
                        marks.append(tok.upper())
            out.append({"name": name, "marks": marks})
    return out


@st.cache_data(show_spinner=False)
def parse_pdf_no_ocr(file_bytes: bytes) -> tuple:
    """Возвращает (students, info). Никакого OCR."""
    students = []
    info = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        info.append(f"Страниц: {len(pdf.pages)}")

        # 1) Пробуем извлечь таблицы по линиям
        all_tables = []
        for page in pdf.pages:
            try:
                tables = page.extract_tables() or []
            except Exception:
                tables = []
            all_tables.extend(tables)
        info.append(f"Найдено таблиц: {len(all_tables)}")

        if all_tables:
            students = _extract_from_tables(all_tables)
            info.append(f"Из таблиц распознано: {len(students)}")

        # 2) Если таблиц нет или мало данных — читаем текст
        if not students:
            text_chunks = []
            for page in pdf.pages:
                try:
                    text_chunks.append(page.extract_text() or "")
                except Exception:
                    text_chunks.append("")
            full_text = "\n".join(text_chunks)
            info.append(f"Текстовых символов: {len(full_text)}")

            students = _extract_from_lines(full_text)
            info.append(f"Из текста распознано: {len(students)}")

            # Первые строки для диагностики
            preview = "\n".join(full_text.splitlines()[:30])
            info.append("Первые 30 строк:\n" + preview)

    return students, info


# ============================================================
# СТРУКТУРА ДАННЫХ И АНАЛИЗ
# ============================================================
def build_rows(students_raw, topics=TOPICS):
    result = []
    for s in students_raw:
        row = {"name": s["name"]}
        marks = s.get("marks", [])
        for i, t in enumerate(topics):
            mark = str(marks[i] if i < len(marks) else "").strip().upper()
            if mark in ("А", "С", "Н"):
                row[t] = "А" if mark == "А" else ("С" if mark == "С" else "")
            else:
                try:
                    v = int(mark)
                    row[t] = v if 1 <= v <= 5 else ""
                except ValueError:
                    row[t] = ""
        result.append(row)
    return result


def analyze(row, topics=TOPICS):
    scores, strong, weak, missing = [], [], [], []
    for t in topics:
        v = row.get(t, "")
        if v in ("А", "С"):
            missing.append({"topic": t, "reason": v})
        elif isinstance(v, int):
            scores.append(v)
            if v >= 4:
                strong.append({"topic": t, "value": v})
            elif v <= 3:
                weak.append({"topic": t, "value": v})
    return {
        "avg": round(sum(scores) / len(scores), 2) if scores else 0.0,
        "strong": strong,
        "weak": weak,
        "missing": missing,
    }


def top_best(rows, n=3, topics=TOPICS):
    return sorted(
        [(r["name"], analyze(r, topics)["avg"]) for r in rows],
        key=lambda x: x[1],
        reverse=True,
    )[:n]


def top_worst(rows, n=3, topics=TOPICS):
    scored = [(r["name"], analyze(r, topics)["avg"]) for r in rows]
    scored = [(nm, a) for nm, a in scored if a > 0]
    return sorted(scored, key=lambda x: x[1])[:n]


def recommendations(row, topics=TOPICS):
    a = analyze(row, topics)
    recs = []
    for w in a["weak"]:
        recs.append(f"📉 «{w['topic']}» — балл {w['value']}. Повторить теорию и решить 5–7 задач.")
    for m in a["missing"]:
        reason = "болел" if m["reason"] == "А" else "уважительная"
        recs.append(f"🩺 «{m['topic']}» — пропуск ({reason}). Изучить тему и ответить на вопросы ниже.")
    if not recs:
        recs.append("✅ Слабых тем и пропусков нет. Отличная работа!")
    return recs


# ============================================================
# UI
# ============================================================
st.markdown("### 👩‍🏫 Учитель")
teacher = st.text_input("ФИО", value="Иванова Айгуль Сериковна", label_visibility="collapsed")
st.markdown(
    f"<div style='background:#e0f2fe;border-left:4px solid #0284c7;"
    f"padding:10px 14px;border-radius:8px;margin-bottom:16px'>"
    f"<b>Учитель:</b> {teacher}</div>",
    unsafe_allow_html=True,
)

st.title("📘 Журнал — Алгебра и начала анализа — 10 А")

st.header("1. Загрузите PDF-журнал (без OCR)")
st.caption("PDF должен содержать текстовый слой (журнал из Excel/Word/электронного дневника). "
           "Если PDF — скан, распознавание не сработает — загрузите Excel.")

uploaded = st.file_uploader("PDF-файл", type=["pdf"])

if uploaded is not None:
    with st.spinner("Читаем PDF…"):
        try:
            students_raw, info = parse_pdf_no_ocr(uploaded.read())
        except Exception as e:
            st.error(f"Ошибка чтения PDF: {e}")
            students_raw, info = [], []

    with st.expander("🩺 Что прочитано из PDF (диагностика)"):
        for line in info:
            st.text(line)

    if students_raw:
        st.session_state["rows"] = build_rows(students_raw)
        st.success(f"✅ Распознано учеников: {len(st.session_state['rows'])}")
    else:
        st.warning(
            "⚠ Ученики не найдены. Возможные причины:\n"
            "1. PDF — скан без текстового слоя (нужен OCR или Excel).\n"
            "2. Нестандартный формат таблицы.\n"
            "3. Откройте диагностику выше и пришлите первые 30 строк."
        )

# Запасной вариант — Excel
with st.expander("📊 Или загрузите готовый Excel (.xlsx)"):
    xlsx = st.file_uploader("Excel", type=["xlsx"], key="xlsx")
    if xlsx is not None:
        try:
            df = pd.read_excel(xlsx)
            if "Ученик" in df.columns:
                df = df.rename(columns={"Ученик": "name"})
            st.session_state["rows"] = df.to_dict(orient="records")
            st.success(f"✅ Загружено из Excel: {len(df)} учеников")
        except Exception as e:
            st.error(f"Ошибка Excel: {e}")

if "rows" not in st.session_state or not st.session_state["rows"]:
    st.info("Загрузите PDF или Excel, чтобы начать анализ.")
    st.stop()

rows = st.session_state["rows"]

# ============================================================
# ТОП ЛУЧШИХ И ХУДШИХ
# ============================================================
st.header("2. Топ лучших и топ худших")
c1, c2 = st.columns(2)
with c1:
    st.subheader("🏆 Топ‑3 лучших")
    for i, (nm, av) in enumerate(top_best(rows), 1):
        st.markdown(f"{['🥇','🥈','🥉'][i-1]} **{nm}** — средний балл **{av:.2f}**")
with c2:
    st.subheader("⚠ Топ‑3 отстающих")
    for i, (nm, av) in enumerate(top_worst(rows), 1):
        st.markdown(f"🔻 **{nm}** — средний балл **{av:.2f}**")

# ============================================================
# СВОДКА
# ============================================================
st.header("3. Сводка по всем ученикам")
summary = []
for r in rows:
    a = analyze(r)
    summary.append({
        "Ученик": r["name"],
        "Средний балл": a["avg"],
        "Сильные темы": ", ".join(f"{x['topic']} ({x['value']})" for x in a["strong"]) or "—",
        "Слабые темы": ", ".join(f"{x['topic']} ({x['value']})" for x in a["weak"]) or "—",
        "Пропуски (А/С)": ", ".join(f"{x['topic']} ({x['reason']})" for x in a["missing"]) or "—",
    })
summary_df = pd.DataFrame(summary).sort_values("Средний балл", ascending=False)
st.dataframe(summary_df, use_container_width=True)

# ============================================================
# ЖУРНАЛ (РЕДАКТИРУЕМЫЙ)
# ============================================================
st.header("4. Данные журнала (можно править)")
journal_df = pd.DataFrame(rows).rename(columns={"name": "Ученик"})
edited = st.data_editor(journal_df, use_container_width=True, num_rows="dynamic")
if st.button("💾 Применить правки"):
    st.session_state["rows"] = edited.rename(columns={"Ученик": "name"}).to_dict(orient="records")
    st.rerun()

# ============================================================
# ГРАФИК ПО УЧЕНИКУ
# ============================================================
st.header("5. График развития по темам")
student_name = st.selectbox("Выберите ученика", [r["name"] for r in rows])
row = next(r for r in rows if r["name"] == student_name)
a = analyze(row)

y_vals, colors = [], []
for t in TOPICS:
    v = row.get(t, "")
    if v in ("А", "С"):
        y_vals.append(None)
        colors.append("#f59e0b")
    elif isinstance(v, int):
        y_vals.append(v)
        colors.append("#16a34a" if v >= 4 else "#dc2626" if v <= 3 else "#2563eb")
    else:
        y_vals.append(None)
        colors.append("#94a3b8")

fig = go.Figure()
fig.add_trace(go.Scatter(
    x=TOPICS, y=y_vals,
    mode="lines+markers",
    line=dict(color="#2563eb", width=2),
    marker=dict(size=10, color=colors),
    name="Балл",
    connectgaps=True,
))
fig.update_layout(
    yaxis=dict(range=[0, 5], dtick=1, title="Балл"),
    xaxis=dict(title="Темы"),
    height=380,
    margin=dict(l=40, r=20, t=30, b=140),
)
st.plotly_chart(fig, use_container_width=True)
st.markdown(
    f"**Средний балл:** {a['avg']:.2f} • "
    f"**Сильных тем:** {len(a['strong'])} • "
    f"**Слабых тем:** {len(a['weak'])} • "
    f"**Пропусков:** {len(a['missing'])}"
)

# ============================================================
# СИЛЬНЫЕ СТОРОНЫ
# ============================================================
st.header("6. Сильные стороны")
if a["strong"]:
    for s in a["strong"]:
        st.markdown(
            f"<div style='background:#ecfdf5;border-left:4px solid #10b981;"
            f"padding:8px 12px;border-radius:8px;margin:4px 0'>"
            f"<b>{s['topic']}</b> — балл <b>{s['value']}</b></div>",
            unsafe_allow_html=True,
        )
else:
    st.info("Пока нет тем с оценкой 4–5.")

# ============================================================
# СЛАБЫЕ + РЕКОМЕНДАЦИИ
# ============================================================
st.header("7. Слабые стороны и рекомендации")
for rec in recommendations(row):
    st.markdown(
        f"<div style='background:#fff7ed;border-left:4px solid #f97316;"
        f"padding:8px 12px;border-radius:8px;margin:4px 0'>{rec}</div>",
        unsafe_allow_html=True,
    )

# ============================================================
# ВОПРОСЫ
# ============================================================
st.header("8. Вопросы для отработки")
targets = list(dict.fromkeys(
    [m["topic"] for m in a["missing"]] + [w["topic"] for w in a["weak"]]
))
if targets:
    for t in targets:
        st.markdown(f"**{t}**")
        for i, q in enumerate(QUESTION_BANK.get(t, []), 1):
            st.markdown(f"{i}. {q}")
else:
    st.info("Нет тем, требующих отработки.")

# ============================================================
# ЭКСПОРТ
# ============================================================
st.header("9. Экспорт и печать")

def build_excel():
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        journal_df.to_excel(writer, sheet_name="Журнал", index=False)
        summary_df.to_excel(writer, sheet_name="Сводка", index=False)
        top_df = pd.DataFrame({
            "Место": [1, 2, 3],
            "Лучшие": [n for n, _ in top_best(rows)],
            "Средний (лучшие)": [a for _, a in top_best(rows)],
            "Худшие": [n for n, _ in top_worst(rows)],
            "Средний (худшие)": [a for _, a in top_worst(rows)],
        })
        top_df.to_excel(writer, sheet_name="Топ", index=False)

        rec_rows = [{"Ученик": r["name"], "Рекомендация": rec}
                    for r in rows for rec in recommendations(r)]
        pd.DataFrame(rec_rows).to_excel(writer, sheet_name="Рекомендации", index=False)

        wb = writer.book
        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="2563EB")
                cell.alignment = Alignment(horizontal="center")
            for col in ws.columns:
                ml = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(ml + 2, 60)
    buffer.seek(0)
    return buffer.getvalue()

st.download_button(
    "⬇ Скачать Excel (журнал + сводка + топ + рекомендации)",
    data=build_excel(),
    file_name="journal_algebra_10A.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
st.markdown("🖨 **Печать:** `Ctrl+P` в браузере.")