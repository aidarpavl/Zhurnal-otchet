import io
import re
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Журнал — Алгебра — 7 Ә", page_icon="📘", layout="wide")

# ============================================================
# БАНК ВОПРОСОВ ПО ТЕМАМ
# ============================================================
QUESTION_BANK = {
    "Натурал көрсеткішті дәреже": [
        "Есептеңіз: 2⁵ · 2³.",
        "Ықшамдаңыз: (a³)⁴ · a².",
        "Есептеңіз: 5⁰ + (−3)².",
        "Ықшамдаңыз: a⁷ : a⁴.",
        "Теңдеуді шешіңіз: 2ˣ = 32.",
    ],
    "Бүтін көрсеткішті дәреже": [
        "Есептеңіз: 2⁻³.",
        "Ықшамдаңыз: (a⁻²)³ · a⁷.",
        "Есептеңіз: (1/3)⁻².",
        "Ықшамдаңыз: a⁻⁵ · a⁸ : a².",
        "Теңдеуді шешіңіз: 3ˣ = 1/27.",
    ],
    "Санның стандарт түрі": [
        "Жазыңыз стандарт түрде: 0,00045.",
        "Жазыңыз стандарт түрде: 12 300 000.",
        "Есептеңіз: (2·10³) · (3·10⁴).",
        "Есептеңіз: (6·10⁵) : (2·10²).",
        "Салыстырыңыз: 3·10⁻² және 4·10⁻³.",
    ],
    "Бірмүшелер және оларға амалдар қолдану": [
        "Ықшамдаңыз: 3a²b · (−2ab³).",
        "Бірмүшенің дәрежесін табыңыз: 5x³y².",
        "Ұқсас мүшелерді біріктіріңіз: 4a + 3a − 2a.",
        "Есептеңіз: (−2x²)³.",
        "Бірмүшені стандарт түрге келтіріңіз: 3·a·4·a².",
    ],
    "Көпмүшелерге амал қолдану": [
        "Қосыңыз: (2x² + 3x − 1) + (x² − 5x + 4).",
        "Азайтыңыз: (5a² − 3a) − (2a² + 4a).",
        "Көбейтіңіз: (x + 2)(x − 3).",
        "Көбейтіңіз: (2a − b)(2a + b).",
        "Ықшамдаңыз: (x + 3)² − x².",
    ],
    "Көпмүшені көбейткіштерге жіктеу": [
        "Жіктеңіз: x² − 9.",
        "Жіктеңіз: 4a² − 25b².",
        "Жіктеңіз: x² + 6x + 9.",
        "Жіктеңіз: 2x² − 8.",
        "Жіктеңіз: a³ − b³.",
    ],
    "Көпмүшені көпмүшеге бөлу": [
        "Бөліңіз: (x² − 4) : (x − 2).",
        "Бөліңіз: (x³ + 1) : (x + 1).",
        "Бөліңіз: (2x² + 5x − 3) : (x + 3).",
        "Бөліңіз: (x³ − 8) : (x − 2).",
        "Бөліңіз: (6x² − x − 2) : (2x + 1).",
    ],
    "Мәтінді есептерді шығару": [
        "Екі санның қосындысы 20, көбейтіндісі 96. Сандарды табыңыз.",
        "Тіктөртбұрыштың периметрі 30 см, ауданы 56 см². Қабырғаларын табыңыз.",
        "Бір сан екіншісінен 3 есе артық. Қосындысы 24. Сандарды табыңыз.",
        "Поезд 60 км/сағ жылдамдықпен 3 сағатта қанша жол жүреді?",
        "Қоспада 20% тұз бар. 300 г қоспада қанша тұз бар?",
    ],
    "Құрамында дәрежесі бар өрнектерді түрлендіру": [
        "Ықшамдаңыз: (a²b)³ · a⁻⁴b.",
        "Ықшамдаңыз: (2x³)² : (4x⁴).",
        "Ықшамдаңыз: (a⁻¹ + b⁻¹) · ab.",
        "Ықшамдаңыз: (x²y⁻³)² · x⁻³y⁴.",
        "Есептеңіз: (2³ · 2⁻⁵) : 2⁻¹.",
    ],
}

DEFAULT_QUESTIONS = [
    "Тақырыптың негізгі анықтамасын тұжырымдаңыз.",
    "Тақырып бойынша типтік есеп шығарыңыз.",
    "Тақырыпта қандай формулалар қолданылады?",
    "Жиі жіберілетін қатені табыңыз және түсіндіріңіз.",
    "Өз бетіңізше есеп құрастырып, шығарыңыз.",
]


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ
# ============================================================
def to_mark(v):
    s = str(v).strip().lower()
    if s in ("а", "н"):
        return "а"
    if not s or s in ("nan", "none"):
        return None
    try:
        n = int(float(s))
        return n if 1 <= n <= 10 else None
    except Exception:
        return None


def norm_topic(t: str) -> str:
    if not t:
        return ""
    s = str(t).strip()
    s = re.sub(r"\s+", " ", s)
    s = s.strip(". ,;")
    return s[:200]


# ============================================================
# ЧТЕНИЕ ЛИСТОВ С ОЦЕНКАМИ
# ============================================================
def read_marks_sheets(xls):
    info = []
    merged = {}

    for sh in xls.sheet_names:
        try:
            raw = pd.read_excel(xls, sheet_name=sh, header=None)
        except Exception as e:
            info.append(f"Лист «{sh}»: ошибка чтения ({e})")
            continue

        if raw.empty:
            continue

        header_row = None
        for i in range(min(len(raw), 10)):
            cell = str(raw.iloc[i, 0]).lower()
            if any(k in cell for k in ["оқушының", "тегі", "ученик", "фио", "ф.и.о"]):
                header_row = i
                break

        if header_row is None:
            continue

        header_dates = raw.iloc[header_row].tolist()[1:]
        data = raw.iloc[header_row + 1:].copy()

        def looks_like_name(v):
            s = str(v).strip()
            if not s or s.lower() in ("nan", "none") or len(s) < 3:
                return False
            if not re.search(r"[А-Яа-яӘәҒғҚқҢңӨөҰұҮүҺһІі]", s):
                return False
            if s.lower().startswith(("бжб", "тоқсан", "макс", "орта", "жиын", "сабақ")):
                return False
            return True

        data = data[data.iloc[:, 0].apply(looks_like_name)]
        if data.empty:
            continue

        added = 0
        for _, r in data.iterrows():
            name = str(r.iloc[0]).strip()
            values = r.tolist()[1:]
            if name not in merged:
                merged[name] = {}
            for date, val in zip(header_dates, values):
                d = str(date).strip()
                if not d or d.lower() in ("nan", "none"):
                    continue
                m = to_mark(val)
                if m is not None:
                    merged[name][d] = m
            added += 1

        info.append(f"Лист оценок «{sh}»: учеников — {added}, дат — {len(header_dates)}")

    return merged, info


# ============================================================
# ЧТЕНИЕ ЛИСТОВ С ТЕМАМИ
# ============================================================
def read_topics_sheets(xls):
    info = []
    date_to_topic = {}

    for sh in xls.sheet_names:
        try:
            raw = pd.read_excel(xls, sheet_name=sh, header=None)
        except Exception:
            continue

        if raw.empty:
            continue

        date_col = None
        topic_col = None
        head_row = None

        for i in range(min(len(raw), 8)):
            row = raw.iloc[i].tolist()
            for j, c in enumerate(row):
                s = str(c).lower()
                if ("күні" in s or "айы" in s) and date_col is None:
                    date_col = j
                    head_row = i
                if "тақырып" in s and topic_col is None:
                    topic_col = j
                    head_row = i
            if date_col is not None and topic_col is not None:
                break

        if date_col is None and topic_col is None:
            continue

        added = 0
        for _, r in raw.iloc[(head_row or 0) + 1:].iterrows():
            date_val = None
            for v in r.tolist():
                s = str(v).strip()
                m = re.match(r"^(\d{1,2})[\.\-/](\d{1,2})(?:[\.\-/](\d{2,4}))?$", s)
                if m:
                    date_val = s
                    break
            if not date_val:
                continue

            topic_val = ""
            if topic_col is not None and topic_col < len(r):
                topic_val = str(r.iloc[topic_col]).strip()
            if not topic_val or topic_val.lower() in ("nan", "none"):
                if topic_col is not None and topic_col + 1 < len(r):
                    topic_val = str(r.iloc[topic_col + 1]).strip()

            topic_val = norm_topic(topic_val)
            if topic_val:
                date_to_topic[date_val] = topic_val
                added += 1

        if added:
            info.append(f"Лист тем «{sh}»: пар «дата → тема» — {added}")

    return date_to_topic, info


# ============================================================
# ЧТЕНИЕ ФИО УЧИТЕЛЯ (лист 2, строка 1)
# ============================================================
def read_teacher_name(xls):
    info = []
    if len(xls.sheet_names) < 2:
        info.append("Листов меньше 2 — ФИО учителя не найдено.")
        return "", info

    second_sheet = xls.sheet_names[1]
    try:
        raw = pd.read_excel(xls, sheet_name=second_sheet, header=None, nrows=3)
    except Exception as e:
        info.append(f"Не удалось прочитать лист «{second_sheet}»: {e}")
        return "", info

    if raw.empty:
        info.append(f"Лист «{second_sheet}» пуст.")
        return "", info

    first_row = raw.iloc[0].tolist()
    for cell in first_row:
        s = str(cell).strip()
        if not s or s.lower() in ("nan", "none"):
            continue
        if re.match(
            r"^([А-ЯЁӘҒҚҢӨҰҮҺІ][а-яёәғқңөұүһі\-]+)(\s+[А-ЯЁӘҒҚҢӨҰҮҺІ][а-яёәғқңөұүһі\-]+){1,3}$",
            s,
        ):
            info.append(f"ФИО учителя найдено на листе «{second_sheet}»: {s}")
            return s, info

    info.append(f"ФИО учителя не найдено в первой строке листа «{second_sheet}».")
    return "", info


# ============================================================
# ЗАГРУЗКА
# ============================================================
@st.cache_data(show_spinner=False)
def read_journal(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))

    marks, info1 = read_marks_sheets(xls)
    date_to_topic, info2 = read_topics_sheets(xls)
    teacher_name, info3 = read_teacher_name(xls)

    info = info1 + info2 + info3

    all_dates = sorted({d for marks in marks.values() for d in marks.keys()},
                       key=lambda x: (len(x), x))

    records = []
    for name, m in marks.items():
        rec = {"Ученик": name}
        for d in all_dates:
            rec[d] = m.get(d, "")
        records.append(rec)

    df = pd.DataFrame(records)

    info.append(
        f"Всего дат: {len(all_dates)}, из них с темами: "
        f"{sum(1 for d in all_dates if d in date_to_topic)}"
    )
    return df, date_to_topic, teacher_name, info


# ============================================================
# АНАЛИЗ
# ============================================================
def analyze_student(row, date_cols, date_to_topic):
    marks = []
    misses = 0
    weak_by_topic = {}
    for d in date_cols:
        v = row.get(d, "")
        topic = date_to_topic.get(d, "")
        if v == "а":
            misses += 1
            continue
        if isinstance(v, int):
            marks.append(v)
            if v <= 5:
                key = topic or f"Дата {d}"
                weak_by_topic.setdefault(key, []).append(v)
    avg = round(sum(marks) / len(marks), 2) if marks else 0.0
    return {
        "avg": avg,
        "count": len(marks),
        "min": min(marks) if marks else 0,
        "max": max(marks) if marks else 0,
        "misses": misses,
        "marks": marks,
        "weak_by_topic": weak_by_topic,
    }


def top_best(df, date_cols, date_to_topic, n=3):
    scored = []
    for _, r in df.iterrows():
        a = analyze_student(r, date_cols, date_to_topic)
        if a["count"] > 0:
            scored.append((r["Ученик"], a["avg"], a["count"]))
    scored.sort(key=lambda x: (x[1], x[2]), reverse=True)
    return scored[:n]


def top_worst(df, date_cols, date_to_topic, n=3):
    scored = []
    for _, r in df.iterrows():
        a = analyze_student(r, date_cols, date_to_topic)
        if a["count"] > 0:
            scored.append((r["Ученик"], a["avg"], a["count"]))
    scored.sort(key=lambda x: (x[1], -x[2]))
    return scored[:n]


def recommendations_for(row, date_cols, date_to_topic):
    a = analyze_student(row, date_cols, date_to_topic)
    recs = []
    if a["count"] == 0:
        recs.append({"type": "warn", "text": "⚠ Нет ни одной оценки. Уточните причину отсутствия."})
        return recs

    if a["avg"] < 5:
        recs.append({"type": "bad", "text":
            f"📉 Средний балл {a['avg']:.2f} — ниже порога. Требуется систематическая работа над ошибками."})
    elif a["avg"] < 7:
        recs.append({"type": "warn", "text":
            f"📊 Средний балл {a['avg']:.2f} — удовлетворительный. Усилить работу над слабыми темами (ниже)."})

    if a["weak_by_topic"]:
        recs.append({"type": "header", "text": "Слабые темы (по датам с низкими оценками):"})
        for topic, marks in a["weak_by_topic"].items():
            avg_topic = sum(marks) / len(marks)
            recs.append({"type": "weak_topic",
                         "topic": topic,
                         "marks": marks,
                         "avg": round(avg_topic, 2)})

    if a["misses"] >= 3:
        recs.append({"type": "warn", "text":
            f"🩺 Пропущено занятий: {a['misses']}. Необходимо отработать пропущенные темы."})

    if not a["weak_by_topic"] and a["avg"] >= 7 and a["misses"] < 3:
        recs.append({"type": "ok", "text":
            "✅ Ученик показывает стабильные результаты. Продолжайте в том же духе!"})

    return recs


def questions_for_topic(topic):
    if topic in QUESTION_BANK:
        return QUESTION_BANK[topic]
    t = topic.lower()
    for key, qs in QUESTION_BANK.items():
        if key.lower() in t or t in key.lower():
            return qs
    if "дәреже" in t and "бүтін" in t:
        return QUESTION_BANK["Бүтін көрсеткішті дәреже"]
    if "дәреже" in t:
        return QUESTION_BANK["Натурал көрсеткішті дәреже"]
    if "стандарт" in t:
        return QUESTION_BANK["Санның стандарт түрі"]
    if "бірмүше" in t:
        return QUESTION_BANK["Бірмүшелер және оларға амалдар қолдану"]
    if "көпмүше" in t and "бөл" in t:
        return QUESTION_BANK["Көпмүшені көпмүшеге бөлу"]
    if "көпмүше" in t and ("жікте" in t or "көбейткіш" in t):
        return QUESTION_BANK["Көпмүшені көбейткіштерге жіктеу"]
    if "көпмүше" in t:
        return QUESTION_BANK["Көпмүшелерге амал қолдану"]
    if "мәтінді" in t or "есептер" in t:
        return QUESTION_BANK["Мәтінді есептерді шығару"]
    return DEFAULT_QUESTIONS


# ============================================================
# UI
# ============================================================
st.markdown("### 👩‍🏫 Учитель")
default_teacher = st.session_state.get("teacher_name", "Жанал Серикхан")
teacher = st.text_input("ФИО", value=default_teacher, label_visibility="collapsed")
st.markdown(
    f"<div style='background:#e0f2fe;border-left:4px solid #0284c7;"
    f"padding:10px 14px;border-radius:8px;margin-bottom:16px'>"
    f"<b>Учитель:</b> {teacher}</div>",
    unsafe_allow_html=True,
)

st.title("📘 Журнал — Алгебра — 7 Ә — 1 четверть")

st.header("1. Загрузите Excel-файл журнала")
st.caption(
    "Оценки — с листов «Алгебра» и «Алгебра (2)». "
    "Темы — с листов «Жанал Серикхан» и «Жанал Серикхан (2)». "
    "ФИО учителя — с листа 2 (первая строка)."
)
uploaded = st.file_uploader("Excel (.xlsx)", type=["xlsx"])

if uploaded is not None:
    with st.spinner("Читаем Excel…"):
        try:
            df, date_to_topic, teacher_name, info = read_journal(uploaded.read())
        except Exception as e:
            st.error(f"Ошибка чтения: {e}")
            df, date_to_topic, teacher_name, info = pd.DataFrame(), {}, "", []

    with st.expander("🩺 Что прочитано из Excel"):
        for line in info:
            st.text(line)

    if df.empty:
        st.warning("⚠ Не удалось найти данные журнала.")
    else:
        st.session_state["df"] = df
        st.session_state["date_to_topic"] = date_to_topic
        st.session_state["teacher_name"] = teacher_name or "Жанал Серикхан"
        st.success(
            f"✅ Загружено учеников: {len(df)}. "
            f"Дат: {len(df.columns) - 1}. "
            f"Тем найдено: {len(date_to_topic)}. "
            f"Учитель: {st.session_state['teacher_name']}."
        )
        st.rerun()

if "df" not in st.session_state or st.session_state["df"].empty:
    st.info("Загрузите Excel, чтобы начать анализ.")
    st.stop()

df = st.session_state["df"]
date_to_topic = st.session_state.get("date_to_topic", {})
date_cols = [c for c in df.columns if c != "Ученик"]

# ============================================================
# ТОП ЛУЧШИХ И ОТСТАЮЩИХ
# ============================================================
st.header("2. Топ лучших и топ отстающих")
best = top_best(df, date_cols, date_to_topic, 3)
worst = top_worst(df, date_cols, date_to_topic, 3)

c1, c2 = st.columns(2)
with c1:
    st.subheader("🏆 Топ‑3 лучших")
    if best:
        for i, (name, avg, cnt) in enumerate(best, 1):
            medal = ["🥇", "🥈", "🥉"][i - 1]
            st.markdown(f"{medal} **{name}** — средний балл **{avg:.2f}** (оценок: {cnt})")
    else:
        st.info("Нет данных")
with c2:
    st.subheader("⚠ Топ‑3 отстающих")
    if worst:
        for i, (name, avg, cnt) in enumerate(worst, 1):
            st.markdown(f"🔻 **{name}** — средний балл **{avg:.2f}** (оценок: {cnt})")
    else:
        st.info("Нет данных")

# ============================================================
# СВОДКА
# ============================================================
st.header("3. Сводка по всем ученикам")
summary = []
for _, r in df.iterrows():
    a = analyze_student(r, date_cols, date_to_topic)
    summary.append({
        "Ученик": r["Ученик"],
        "Средний балл": a["avg"],
        "Оценок": a["count"],
        "Минимум": a["min"],
        "Максимум": a["max"],
        "Пропусков (а)": a["misses"],
        "Слабых тем": len(a["weak_by_topic"]),
    })
summary_df = pd.DataFrame(summary).sort_values("Средний балл", ascending=False)
st.dataframe(summary_df, use_container_width=True)

# ============================================================
# ВСЕ ОЦЕНКИ
# ============================================================
st.header("4. Все оценки учеников по датам")
st.caption("«а» — отсутствовал. Значения — оценки 1–10.")
st.dataframe(df, use_container_width=True)

# ============================================================
# ТЕМЫ ПО ДАТАМ
# ============================================================
with st.expander("📚 Темы по датам (из листов «Жанал Серикхан»)"):
    if date_to_topic:
        topic_df = pd.DataFrame(
            [{"Дата": d, "Тема": t} for d, t in date_to_topic.items()]
        )
        st.dataframe(topic_df, use_container_width=True)
    else:
        st.info("Темы не найдены. Проверьте листы «Жанал Серикхан» и «Жанал Серикхан (2)».")

# ============================================================
# ГРАФИК
# ============================================================
st.header("5. График развития выбранного ученика")
student_name = st.selectbox("Ученик", df["Ученик"].tolist())
row = df[df["Ученик"] == student_name].iloc[0]
a = analyze_student(row, date_cols, date_to_topic)

x_vals, y_vals, colors = [], [], []
for d in date_cols:
    v = row.get(d, "")
    x_vals.append(d)
    if v == "а":
        y_vals.append(None)
        colors.append("#f59e0b")
    elif isinstance(v, int):
        y_vals.append(v)
        colors.append("#16a34a" if v >= 7 else "#dc2626" if v <= 5 else "#2563eb")
    else:
        y_vals.append(None)
        colors.append("#94a3b8")

fig = go.Figure()
fig.add_trace(go.Scatter(
    x=x_vals, y=y_vals,
    mode="lines+markers",
    line=dict(color="#2563eb", width=2),
    marker=dict(size=10, color=colors),
    name="Балл",
    connectgaps=True,
))
fig.update_layout(
    yaxis=dict(range=[0, 10], dtick=1, title="Балл"),
    xaxis=dict(title="Даты"),
    height=380,
    margin=dict(l=40, r=20, t=30, b=140),
)
st.plotly_chart(fig, use_container_width=True)

st.markdown(
    f"**Средний балл:** {a['avg']:.2f} • "
    f"**Оценок:** {a['count']} • "
    f"**Минимум:** {a['min']} • "
    f"**Максимум:** {a['max']} • "
    f"**Пропусков (а):** {a['misses']} • "
    f"**Слабых тем:** {len(a['weak_by_topic'])}"
)

# ============================================================
# РЕКОМЕНДАЦИИ
# ============================================================
st.header("6. Рекомендации по темам")
recs = recommendations_for(row, date_cols, date_to_topic)
for rec in recs:
    if rec["type"] == "header":
        st.markdown(f"**{rec['text']}**")
    elif rec["type"] == "weak_topic":
        st.markdown(
            f"<div style='background:#fff7ed;border-left:4px solid #f97316;"
            f"padding:10px 14px;border-radius:8px;margin:8px 0'>"
            f"<b>Тема:</b> {rec['topic']}<br>"
            f"<b>Оценки:</b> {', '.join(str(m) for m in rec['marks'])} "
            f"(средний по теме: {rec['avg']})<br>"
            f"<b>Рекомендация:</b> повторить теорию, разобрать типовые задачи, "
            f"выполнить 5–7 упражнений по теме.</div>",
            unsafe_allow_html=True,
        )
        qs = questions_for_topic(rec["topic"])
        st.markdown(f"**Вопросы по теме «{rec['topic']}»:**")
        for i, q in enumerate(qs, 1):
            st.markdown(f"{i}. {q}")
    elif rec["type"] == "ok":
        st.markdown(
            f"<div style='background:#ecfdf5;border-left:4px solid #10b981;"
            f"padding:10px 14px;border-radius:8px;margin:8px 0'>{rec['text']}</div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"<div style='background:#fff7ed;border-left:4px solid #f97316;"
            f"padding:10px 14px;border-radius:8px;margin:8px 0'>{rec['text']}</div>",
            unsafe_allow_html=True,
        )

# ============================================================
# ЭКСПОРТ
# ============================================================
st.header("7. Экспорт и печать")


def build_excel():
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Все оценки", index=False)
        summary_df.to_excel(writer, sheet_name="Сводка", index=False)

        n = max(len(best), len(worst), 1)

        def pad(lst, size):
            lst = list(lst)
            return lst + [""] * (size - len(lst))

        top_df = pd.DataFrame({
            "Место": list(range(1, n + 1)),
            "Лучшие": pad([x[0] for x in best], n),
            "Средний балл (лучшие)": pad([x[1] for x in best], n),
            "Отстающие": pad([x[0] for x in worst], n),
            "Средний балл (отстающие)": pad([x[1] for x in worst], n),
        })
        top_df.to_excel(writer, sheet_name="Топ", index=False)

        if date_to_topic:
            topics_df = pd.DataFrame(
                [{"Дата": d, "Тема": t} for d, t in date_to_topic.items()]
            )
            topics_df.to_excel(writer, sheet_name="Темы", index=False)

        rec_rows = []
        q_rows = []
        for _, r in df.iterrows():
            for rec in recommendations_for(r, date_cols, date_to_topic):
                if rec["type"] == "weak_topic":
                    rec_rows.append({
                        "Ученик": r["Ученик"],
                        "Тема": rec["topic"],
                        "Оценки": ", ".join(str(m) for m in rec["marks"]),
                        "Средний по теме": rec["avg"],
                        "Рекомендация": "Повторить теорию и решить 5–7 задач",
                    })
                    for q in questions_for_topic(rec["topic"]):
                        q_rows.append({
                            "Ученик": r["Ученик"],
                            "Тема": rec["topic"],
                            "Вопрос": q,
                        })
                else:
                    rec_rows.append({
                        "Ученик": r["Ученик"],
                        "Тема": "",
                        "Оценки": "",
                        "Средний по теме": "",
                        "Рекомендация": rec["text"],
                    })
        if not rec_rows:
            rec_rows = [{"Ученик": "—", "Тема": "", "Оценки": "",
                         "Средний по теме": "", "Рекомендация": "Нет данных"}]
        pd.DataFrame(rec_rows).to_excel(writer, sheet_name="Рекомендации", index=False)
        if q_rows:
            pd.DataFrame(q_rows).to_excel(writer, sheet_name="Вопросы", index=False)

        wb = writer.book
        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="2563EB")
                cell.alignment = Alignment(horizontal="center")
            for col in ws.columns:
                ml = max((len(str(c.value)) if c.value else 0) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(ml + 2, 60)

    buffer.seek(0)
    return buffer.getvalue()


try:
    excel_bytes = build_excel()
    st.download_button(
        "⬇ Скачать Excel (оценки + сводка + топ + темы + рекомендации + вопросы)",
        data=excel_bytes,
        file_name="journal_algebra_7A.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
except Exception as e:
    st.error(f"Ошибка формирования Excel: {e}")

st.markdown("🖨 **Печать:** `Ctrl+P` в браузере.")
