import io
import re
import traceback
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Журнал", page_icon="📘", layout="wide")

# ============================================================
# ВСПОМОГАТЕЛЬНОЕ
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


# ============================================================
# ЧТЕНИЕ EXCEL
# ============================================================
@st.cache_data(show_spinner=False)
def read_journal(file_bytes: bytes):
    info = []
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    info.append(f"Листов в файле: {len(xls.sheet_names)} — {xls.sheet_names}")

    journal_frames = []

    for sh in xls.sheet_names:
        try:
            raw = pd.read_excel(xls, sheet_name=sh, header=None)
        except Exception as e:
            info.append(f"Лист «{sh}»: ОШИБКА чтения — {type(e).__name__}: {e}")
            continue

        if raw.empty:
            info.append(f"Лист «{sh}»: пустой")
            continue

        header_row = None
        for i in range(min(len(raw), 10)):
            cell = str(raw.iloc[i, 0]).lower()
            if any(k in cell for k in ["оқушының", "тегі", "ученик", "ф.и.о", "фио", "аты-жөні"]):
                header_row = i
                break

        if header_row is None:
            info.append(f"Лист «{sh}»: заголовок журнала не найден")
            continue

        header_dates = raw.iloc[header_row].tolist()[1:]
        data = raw.iloc[header_row + 1:].copy()

        def looks_like_name(v):
            s = str(v).strip()
            if not s or s.lower() in ("nan", "none") or len(s) < 3:
                return False
            if not re.search(r"[А-Яа-яӘәҒғҚқҢңӨөҰұҮүҺһІі]", s):
                return False
            if s.lower().startswith(("бжб", "тоқсан", "макс", "орта", "жиын")):
                return False
            return True

        data = data[data.iloc[:, 0].apply(looks_like_name)]

        if data.empty:
            info.append(f"Лист «{sh}»: строки с ФИО не найдены")
            continue

        rows = []
        for _, r in data.iterrows():
            name = str(r.iloc[0]).strip()
            values = r.tolist()[1:]
            marks = {}
            for date, val in zip(header_dates, values):
                d = str(date).strip()
                if not d or d.lower() in ("nan", "none"):
                    continue
                m = to_mark(val)
                if m is not None:
                    marks[d] = m
            rows.append({"Ученик": name, "marks": marks})

        journal_frames.append(rows)
        info.append(f"Лист «{sh}»: учеников — {len(rows)}, дат — {len(header_dates)} ✓")

    if not journal_frames:
        return pd.DataFrame(), "", info

    merged = {}
    for block in journal_frames:
        for row in block:
            name = row["Ученик"]
            if name not in merged:
                merged[name] = {}
            merged[name].update(row["marks"])

    all_dates = sorted({d for m in merged.values() for d in m.keys()})

    records = []
    for name, marks in merged.items():
        rec = {"Ученик": name}
        for d in all_dates:
            rec[d] = marks.get(d, "")
        records.append(rec)

    df = pd.DataFrame(records)

    teacher_name = ""
    if len(xls.sheet_names) >= 2:
        second = xls.sheet_names[1]
        try:
            raw2 = pd.read_excel(xls, sheet_name=second, header=None, nrows=3)
            if not raw2.empty:
                for cell in raw2.iloc[0].tolist():
                    s = str(cell).strip()
                    if not s or s.lower() in ("nan", "none"):
                        continue
                    if re.match(
                        r"^([А-ЯЁӘҒҚҢӨҰҮҺІ][а-яёәғқңөұүһі\-]+)"
                        r"(\s+[А-ЯЁӘҒҚҢӨҰҮҺІ][а-яёәғқңөұүһі\-]+){1,3}$",
                        s,
                    ):
                        teacher_name = s
                        info.append(f"ФИО учителя: «{s}» ✓")
                        break
        except Exception as e:
            info.append(f"Ошибка чтения листа «{second}»: {e}")

    info.append(f"Итог: учеников — {len(df)}, дат — {len(all_dates)}")
    return df, teacher_name, info


# ============================================================
# АНАЛИЗ
# ============================================================
def analyze_student(row, date_cols):
    marks, misses = [], 0
    for d in date_cols:
        v = row.get(d, "")
        if v == "а":
            misses += 1
        elif isinstance(v, int):
            marks.append(v)
    avg = round(sum(marks) / len(marks), 2) if marks else 0.0
    total = len(marks)
    quality = sum(1 for m in marks if m >= 7)
    progress = sum(1 for m in marks if m >= 4)
    return {
        "avg": avg, "count": total,
        "min": min(marks) if marks else 0,
        "max": max(marks) if marks else 0,
        "misses": misses, "marks": marks,
        "quality_pct": round(quality / total * 100, 1) if total else 0.0,
        "progress_pct": round(progress / total * 100, 1) if total else 0.0,
    }


def top_best(df, date_cols, n=3):
    scored = []
    for _, r in df.iterrows():
        a = analyze_student(r, date_cols)
        if a["count"] > 0:
            scored.append((r["Ученик"], a["avg"], a["count"]))
    scored.sort(key=lambda x: (x[1], x[2]), reverse=True)
    return scored[:n]


def top_worst(df, date_cols, n=3):
    scored = []
    for _, r in df.iterrows():
        a = analyze_student(r, date_cols)
        if a["count"] > 0:
            scored.append((r["Ученик"], a["avg"], a["count"]))
    scored.sort(key=lambda x: (x[1], -x[2]))
    return scored[:n]


def recommendations_for(row, date_cols):
    a = analyze_student(row, date_cols)
    recs = []
    if a["count"] == 0:
        recs.append("⚠ Нет ни одной оценки.")
        return recs
    if a["avg"] < 5:
        recs.append(f"📉 Средний балл {a['avg']:.2f} — ниже порога.")
    elif a["avg"] < 7:
        recs.append(f"📊 Средний балл {a['avg']:.2f} — удовлетворительный.")
    if a["misses"] >= 3:
        recs.append(f"🩺 Пропущено занятий: {a['misses']}.")
    low = [m for m in a["marks"] if m <= 5]
    if low:
        recs.append(f"🔻 Низких оценок (≤5): {len(low)}.")
    if not recs:
        recs.append("✅ Ученик показывает стабильные результаты.")
    return recs


def color_for_pct(pct):
    if pct >= 75:
        return "#16a34a"
    if pct >= 50:
        return "#f59e0b"
    return "#dc2626"


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
    "Оценки — с листов «Алгебра» / «Алгебра (2)». "
    "ФИО учителя — с листа 2. "
    "Все листы с журналом объединяются."
)

if "df" in st.session_state and not st.session_state["df"].empty:
    col_a, col_b = st.columns([3, 1])
    with col_b:
        if st.button("🔄 Загрузить другой файл", use_container_width=True):
            for k in ["df", "teacher_name", "info"]:
                st.session_state.pop(k, None)
            st.rerun()
    with col_a:
        st.success(
            f"✅ Файл загружен. Учеников: {len(st.session_state['df'])}, "
            f"дат: {len(st.session_state['df'].columns) - 1}."
        )

if "df" not in st.session_state or st.session_state["df"].empty:
    uploaded = st.file_uploader("Excel (.xlsx)", type=["xlsx", "xls"])

    if uploaded is not None:
        st.info(
            f"📦 **Файл:** `{uploaded.name}`  \n"
            f"📏 **Размер:** {round(uploaded.size / 1024, 1)} КБ"
        )
        with st.spinner("Читаем Excel…"):
            try:
                data = uploaded.read()
                df_new, teacher_new, info_new = read_journal(data)
            except Exception as e:
                st.error(f"❌ **{type(e).__name__}:** {e}")
                st.code(traceback.format_exc())
                df_new, teacher_new, info_new = pd.DataFrame(), "", []

        with st.expander("🩺 Диагностика чтения Excel", expanded=True):
            for line in (info_new or ["(нет данных)"]):
                st.text(line)

        if df_new.empty:
            st.warning(
                "⚠ Данные журнала не найдены. Проверьте, что в первой колонке "
                "листа «Алгебра» есть слово «Оқушының» / «Ученик» / «ФИО»."
            )
        else:
            st.session_state["df"] = df_new
            st.session_state["teacher_name"] = teacher_new or "Жанал Серикхан"
            st.session_state["info"] = info_new
            st.rerun()
    else:
        st.info("Загрузите Excel, чтобы начать анализ.")
        st.stop()

# ============================================================
# ОСНОВНОЙ АНАЛИЗ
# ============================================================
df = st.session_state["df"]
date_cols = [c for c in df.columns if c != "Ученик"]

with st.expander("🩺 Что прочитано из Excel"):
    for line in st.session_state.get("info", []):
        st.text(line)

# ---------- ТОП ----------
st.header("2. Топ лучших и топ отстающих")
best = top_best(df, date_cols, 3)
worst = top_worst(df, date_cols, 3)

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

# ---------- СВОДКА ----------
st.header("3. Сводка по всем ученикам")
summary = []
for _, r in df.iterrows():
    a = analyze_student(r, date_cols)
    summary.append({
        "Ученик": r["Ученик"],
        "Средний балл": a["avg"],
        "Білім сапасы %": a["quality_pct"],
        "Үлгерім %": a["progress_pct"],
        "Оценок": a["count"],
        "Минимум": a["min"],
        "Максимум": a["max"],
        "Пропусков (а)": a["misses"],
        "Низких (≤5)": len([m for m in a["marks"] if m <= 5]),
    })
summary_df = pd.DataFrame(summary).sort_values("Средний балл", ascending=False)
st.dataframe(summary_df, use_container_width=True)

# ---------- ВСЕ ОЦЕНКИ ----------
st.header("4. Все оценки учеников по датам")
st.caption("«а» — отсутствовал. Оценки — 1–10.")
st.dataframe(df, use_container_width=True)

# ---------- ГРАФИК ----------
st.header("5. График развития выбранного ученика")
student_name = st.selectbox("Ученик", df["Ученик"].tolist())
row = df[df["Ученик"] == student_name].iloc[0]
a = analyze_student(row, date_cols)

y_vals, colors, x_vals = [], [], []
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

q_pct = a["quality_pct"]
p_pct = a["progress_pct"]
st.markdown(
    f"""
    <div style='display:flex;gap:12px;flex-wrap:wrap;margin:12px 0'>
      <div style='flex:1;min-width:200px;background:#f0f9ff;border-left:4px solid {color_for_pct(q_pct)};
                  padding:12px 16px;border-radius:8px'>
        <div style='font-size:12px;color:#475569'>Білім сапасы</div>
        <div style='font-size:26px;font-weight:700;color:{color_for_pct(q_pct)}'>{q_pct:.1f}%</div>
        <div style='font-size:12px;color:#64748b'>доля оценок 7–10</div>
      </div>
      <div style='flex:1;min-width:200px;background:#f0f9ff;border-left:4px solid {color_for_pct(p_pct)};
                  padding:12px 16px;border-radius:8px'>
        <div style='font-size:12px;color:#475569'>Үлгерім</div>
        <div style='font-size:26px;font-weight:700;color:{color_for_pct(p_pct)}'>{p_pct:.1f}%</div>
        <div style='font-size:12px;color:#64748b'>доля оценок ≥ 4</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"**Средний балл:** {a['avg']:.2f} • "
    f"**Оценок:** {a['count']} • "
    f"**Минимум:** {a['min']} • "
    f"**Максимум:** {a['max']} • "
    f"**Пропусков (а):** {a['misses']}"
)

# ---------- РЕКОМЕНДАЦИИ ----------
st.header("6. Рекомендации")
for rec in recommendations_for(row, date_cols):
    st.markdown(
        f"<div style='background:#fff7ed;border-left:4px solid #f97316;"
        f"padding:8px 12px;border-radius:8px;margin:4px 0'>{rec}</div>",
        unsafe_allow_html=True,
    )

# ============================================================
# АНАЛИТИЧЕСКИЙ ОТЧЁТ ПО КЛАССУ
# ============================================================
st.header("7. Аналитический отчёт по классу")

# --- Общие показатели ---
class_avgs = [analyze_student(r, date_cols)["avg"] for _, r in df.iterrows()]
class_avgs = [x for x in class_avgs if x > 0]
class_avg = round(sum(class_avgs) / len(class_avgs), 2) if class_avgs else 0.0

quality_list = [analyze_student(r, date_cols)["quality_pct"] for _, r in df.iterrows()]
progress_list = [analyze_student(r, date_cols)["progress_pct"] for _, r in df.iterrows()]
class_quality = round(sum(quality_list) / len(quality_list), 1) if quality_list else 0.0
class_progress = round(sum(progress_list) / len(progress_list), 1) if progress_list else 0.0

total_misses = sum(analyze_student(r, date_cols)["misses"] for _, r in df.iterrows())
low_marks_total = sum(
    len([m for m in analyze_student(r, date_cols)["marks"] if m <= 5])
    for _, r in df.iterrows()
)
total_marks = sum(analyze_student(r, date_cols)["count"] for _, r in df.iterrows())

st.markdown(
    f"""
**Класс:** {teacher} — 7 Ә — 1 четверть
**Учеников:** {len(df)}
**Всего выставленных оценок:** {total_marks}
**Пропусков (а):** {total_misses}
**Средний балл по классу:** {class_avg}
**Білім сапасы (класс):** {class_quality}%
**Үлгерім (класс):** {class_progress}%
**Низких оценок (≤5):** {low_marks_total}
"""
)

# --- Разбивка по уровням ---
st.subheader("Распределение по уровням успеваемости")

levels = {"7–10 (жоғары)": 0, "4–6 (орта)": 0, "1–3 (төмен)": 0}
for _, r in df.iterrows():
    a = analyze_student(r, date_cols)
    if not a["marks"]:
        continue
    if a["avg"] >= 7:
        levels["7–10 (жоғары)"] += 1
    elif a["avg"] >= 4:
        levels["4–6 (орта)"] += 1
    else:
        levels["1–3 (төмен)"] += 1

for k, v in levels.items():
    st.markdown(f"- **{k}:** {v} учеников")

# --- Топ и отстающие ещё раз для отчёта ---
st.subheader("Рейтинг")
if best:
    st.markdown("**Лучшие:**")
    for i, (name, avg, cnt) in enumerate(best, 1):
        st.markdown(f"{i}. {name} — средний балл {avg:.2f} (оценок: {cnt})")

if worst:
    st.markdown("**Требуют внимания:**")
    for i, (name, avg, cnt) in enumerate(worst, 1):
        st.markdown(f"{i}. {name} — средний балл {avg:.2f} (оценок: {cnt})")

# --- Ученики с низкими показателями ---
st.subheader("Ученики с низкими показателями")
low_students = []
for _, r in df.iterrows():
    a = analyze_student(r, date_cols)
    if a["count"] > 0 and (a["avg"] < 5 or a["quality_pct"] < 50 or a["misses"] >= 3):
        low_students.append((r["Ученик"], a["avg"], a["quality_pct"], a["progress_pct"], a["misses"]))

if low_students:
    low_df = pd.DataFrame(
        low_students,
        columns=["Ученик", "Средний балл", "Білім сапасы %", "Үлгерім %", "Пропусков"],
    )
    st.dataframe(low_df, use_container_width=True)
else:
    st.success("Учеников с критически низкими показателями нет.")

# --- Рекомендации по классу ---
st.subheader("Рекомендации по классу")
class_recs = []
if class_quality < 50:
    class_recs.append("📉 Білім сапасы ниже 50% — рекомендуется усилить работу над темами с низкими оценками.")
if class_progress < 90:
    class_recs.append("⚠ Үлгерім ниже 90% — обратите внимание на неуспевающих учеников.")
if total_misses > len(df) * 2:
    class_recs.append("🩺 Много пропусков — отработать пропущенные темы индивидуально.")
if low_marks_total > total_marks * 0.2:
    class_recs.append("🔻 Более 20% низких оценок — провести повторение материала.")
if not class_recs:
    class_recs.append("✅ Показатели класса в норме. Продолжайте в том же духе.")

for rec in class_recs:
    st.markdown(
        f"<div style='background:#fff7ed;border-left:4px solid #f97316;"
        f"padding:8px 12px;border-radius:8px;margin:4px 0'>{rec}</div>",
        unsafe_allow_html=True,
    )

# ============================================================
# ПРОМПТ ДЛЯ ДАЛЬНЕЙШИХ РАБОТ
# ============================================================
st.header("8. Готовый промпт для дальнейших работ")

st.caption(
    "Скопируйте текст ниже и вставьте его в новую сессию при работе с ИИ — "
    "это даст полный контекст проекта и ускорит доработку."
)

prompt_text = f"""Проект: Streamlit-приложение «Журнал — Алгебра — 7 Ә — 1 четверть».

Назначение:
- загрузка Excel-файла журнала BilimClass (несколько листов);
- чтение оценок с листов типа «Алгебра», «Алгебра (2)»;
- чтение тем — с листов «Жанал Серикхан», «Жанал Серикхан (2)»;
- чтение ФИО учителя — с листа № 2, первая строка;
- расчёт среднего балла, Білім сапасы (%) и Үлгерім (%) по каждому ученику;
- график развития по датам (Plotly);
- топ‑3 лучших и топ‑3 отстающих;
- рекомендации по слабым оценкам и пропускам («а»);
- аналитический отчёт по классу;
- экспорт в Excel: Все оценки, Сводка, Топ, Рекомендации;
- печать через Ctrl+P.

Текущий файл: журнал за 1 четверть, оценки 1–10, буква «а» — отсутствовал.
Текущий учитель: {teacher}
Всего учеников в файле: {len(df)}
Всего дат (столбцов): {len(date_cols)}

Что нужно доработать:
1. Связать даты с темами из листов «Жанал Серикхан» и выдавать рекомендации по конкретным темам.
2. По каждой слабой теме — 5 вопросов для отработки (банк уже частично есть).
3. Сохранение отчёта в PDF с шапкой школы.
4. Возможность загрузить несколько четвертей и показать годовую динамику.
5. Автоматическая рассылка отчёта родителям (по email из отдельного файла).

Требования к коду:
- Python 3.11+, Streamlit, pandas, openpyxl, plotly;
- без OCR (только текстовый слой Excel);
- стиль кода — с русскими комментариями и разбивкой на функции;
- не удалять существующую функциональность, только добавлять.
"""

st.code(prompt_text, language="text")

st.download_button(
    "⬇ Скачать промпт (.txt)",
    data=prompt_text.encode("utf-8"),
    file_name="prompt_journal.txt",
    mime="text/plain",
)

# ============================================================
# ЭКСПОРТ EXCEL
# ============================================================
st.header("9. Экспорт и печать")


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

        rec_rows = []
        for _, r in df.iterrows():
            for rec in recommendations_for(r, date_cols):
                rec_rows.append({"Ученик": r["Ученик"], "Рекомендация": rec})
        if not rec_rows:
            rec_rows = [{"Ученик": "—", "Рекомендация": "Нет данных"}]
        pd.DataFrame(rec_rows).to_excel(writer, sheet_name="Рекомендации", index=False)

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
        "⬇ Скачать Excel (оценки + сводка + топ + рекомендации)",
        data=excel_bytes,
        file_name="journal_algebra_7A.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
except Exception as e:
    st.error(f"Ошибка формирования Excel: {e}")

st.markdown("🖨 **Печать:** `Ctrl+P` в браузере.")
