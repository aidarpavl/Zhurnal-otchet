import io
import re
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Журнал", page_icon="📘", layout="wide")

# ============================================================
# ВСПОМОГАТЕЛЬНОЕ
# ============================================================
def is_mark(v) -> bool:
    """Проверяет, является ли значение оценкой 1–10."""
    s = str(v).strip().lower()
    if s in ("а", "н", ""):
        return False
    try:
        n = int(float(s))
        return 1 <= n <= 10
    except Exception:
        return False


def to_mark(v):
    """Преобразует значение в int или 'а' (отсутствовал) или None."""
    s = str(v).strip().lower()
    if s in ("а", "н"):
        return "а"
    if not s:
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
    """
    Читает все листы Excel, ищет листы с журналом (первый столбец — ФИО)
    и склеивает их по вертикали.
    Дополнительно читает ФИО учителя с листа № 2 (первая строка).
    Возвращает: (df, teacher_name, sheet_info)
    """
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_names = xls.sheet_names

    journal_frames = []
    sheet_info = []

    for sh in sheet_names:
        try:
            raw = pd.read_excel(xls, sheet_name=sh, header=None)
        except Exception as e:
            sheet_info.append(f"Лист «{sh}»: ошибка чтения ({e})")
            continue

        if raw.empty:
            sheet_info.append(f"Лист «{sh}»: пустой")
            continue

        # Ищем строку заголовка: где в первой ячейке есть "Оқушының"/"тегі"/"Ученик"
        header_row = None
        for i in range(min(len(raw), 10)):
            cell = str(raw.iloc[i, 0]).lower()
            if any(k in cell for k in ["оқушының", "тегі", "ученик", "ф.и.о", "фио", "аты-жөні"]):
                header_row = i
                break

        if header_row is None:
            sheet_info.append(f"Лист «{sh}»: не найден заголовок журнала")
            continue

        header_dates = raw.iloc[header_row].tolist()[1:]
        data = raw.iloc[header_row + 1:].copy()

        def looks_like_name(v):
            s = str(v).strip()
            if not s or s.lower() in ("nan", "none"):
                return False
            if len(s) < 3:
                return False
            if not re.search(r"[А-Яа-яӘәҒғҚқҢңӨөҰұҮүҺһІі]", s):
                return False
            if s.lower().startswith(("бжб", "тоқсан", "макс", "орта", "жиын")):
                return False
            return True

        data = data[data.iloc[:, 0].apply(looks_like_name)]

        if data.empty:
            sheet_info.append(f"Лист «{sh}»: нет учеников")
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
        sheet_info.append(f"Лист «{sh}»: учеников — {len(rows)}, столбцов‑дат — {len(header_dates)}")

    if not journal_frames:
        return pd.DataFrame(), "", sheet_info

    # Склеиваем всех учеников по имени
    merged = {}
    for block in journal_frames:
        for row in block:
            name = row["Ученик"]
            if name not in merged:
                merged[name] = {}
            merged[name].update(row["marks"])

    all_dates = sorted({d for marks in merged.values() for d in marks.keys()})

    records = []
    for name, marks in merged.items():
        rec = {"Ученик": name}
        for d in all_dates:
            rec[d] = marks.get(d, "")
        records.append(rec)

    df = pd.DataFrame(records)

    # ---------- ФИО учителя с листа № 2 ----------
    teacher_name = ""
    if len(sheet_names) >= 2:
        second_sheet = sheet_names[1]
        try:
            raw2 = pd.read_excel(xls, sheet_name=second_sheet, header=None, nrows=3)
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
                        sheet_info.append(f"ФИО учителя найдено на листе «{second_sheet}»: {s}")
                        break
            if not teacher_name:
                sheet_info.append(f"ФИО учителя не найдено в первой строке листа «{second_sheet}».")
        except Exception as e:
            sheet_info.append(f"Не удалось прочитать лист «{second_sheet}»: {e}")

    return df, teacher_name, sheet_info


# ============================================================
# АНАЛИЗ
# ============================================================
def analyze_student(row, date_cols):
    """Возвращает avg, count, min, max, misses, marks, quality_pct, progress_pct."""
    marks = []
    misses = 0
    for d in date_cols:
        v = row.get(d, "")
        if v == "а":
            misses += 1
        elif isinstance(v, int):
            marks.append(v)

    avg = round(sum(marks) / len(marks), 2) if marks else 0.0
    total = len(marks)

    # Білім сапасы — доля оценок 7–10
    quality = sum(1 for m in marks if m >= 7)
    quality_pct = round(quality / total * 100, 1) if total else 0.0

    # Үлгерім — доля положительных оценок (>= 4)
    progress = sum(1 for m in marks if m >= 4)
    progress_pct = round(progress / total * 100, 1) if total else 0.0

    return {
        "avg": avg,
        "count": total,
        "min": min(marks) if marks else 0,
        "max": max(marks) if marks else 0,
        "misses": misses,
        "marks": marks,
        "quality_pct": quality_pct,
        "progress_pct": progress_pct,
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
        recs.append("⚠ Нет ни одной оценки. Уточните причину отсутствия.")
        return recs
    if a["avg"] < 5:
        recs.append(
            f"📉 Средний балл {a['avg']:.2f} — ниже порога. "
            f"Рекомендуется повторить текущие темы и выполнить дополнительные задания."
        )
    elif a["avg"] < 7:
        recs.append(
            f"📊 Средний балл {a['avg']:.2f} — удовлетворительный. "
            f"Стоит усилить работу над темами с низкими оценками."
        )
    if a["misses"] >= 3:
        recs.append(
            f"🩺 Пропущено занятий: {a['misses']}. "
            f"Необходимо отработать пропущенные темы по учебнику и ответить на вопросы."
        )
    low = [m for m in a["marks"] if m <= 5]
    if low:
        recs.append(
            f"🔻 Низких оценок (≤5): {len(low)}. "
            f"Рекомендуется повторить материал и сдать задолженности."
        )
    if not recs:
        recs.append("✅ Ученик показывает стабильные результаты. Продолжайте в том же духе!")
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

# ---------- ЗАГРУЗКА EXCEL ----------
st.header("1. Загрузите Excel-файл журнала")
st.caption(
    "Подойдёт файл BilimClass с листами «Алгебра», «Алгебра (2)» и т.п. "
    "Все листы с журналом объединяются. ФИО учителя берётся с листа 2."
)
uploaded = st.file_uploader("Excel (.xlsx)", type=["xlsx"])

if uploaded is not None:
    st.write(f"Размер файла: {round(uploaded.size / 1024, 1)} КБ")
    with st.spinner("Читаем Excel…"):
        try:
            df, teacher_name, info = read_journal(uploaded.read())
        except Exception as e:
            import traceback
            st.error(f"Ошибка чтения: {type(e).__name__}: {e}")
            st.code(traceback.format_exc())
            df, teacher_name, info = pd.DataFrame(), "", []

    with st.expander("🩺 Что прочитано из Excel"):
        for line in info:
            st.text(line)

    if df.empty:
        st.warning("⚠ Не удалось найти данные журнала. Проверьте, что в файле есть лист с ФИО учеников.")
    else:
        st.session_state["df"] = df
        st.session_state["teacher_name"] = teacher_name or "Жанал Серикхан"
        st.success(f"✅ Загружено учеников: {len(df)}. Столбцов (дат): {len(df.columns) - 1}.")
        st.rerun()

if "df" not in st.session_state or st.session_state["df"].empty:
    st.info("Загрузите Excel, чтобы начать анализ.")
    st.stop()

df = st.session_state["df"]
date_cols = [c for c in df.columns if c != "Ученик"]

# ============================================================
# ТОП ЛУЧШИХ И ХУДШИХ
# ============================================================
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

# ============================================================
# СВОДКА ПО ВСЕМ УЧЕНИКАМ
# ============================================================
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

# ============================================================
# ВСЕ ОЦЕНКИ УЧЕНИКОВ
# ============================================================
st.header("4. Все оценки учеников по датам")
st.caption("Показаны все оценки, которые есть в Excel. «а» — отсутствовал.")
st.dataframe(df, use_container_width=True)

# ============================================================
# ГРАФИК ПО УЧЕНИКУ
# ============================================================
st.header("5. График развития выбранного ученика")
student_name = st.selectbox("Ученик", df["Ученик"].tolist())
row = df[df["Ученик"] == student_name].iloc[0]
a = analyze_student(row, date_cols)

# --- График ---
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

# --- Плашки с процентами ---
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

# ============================================================
# РЕКОМЕНДАЦИИ
# ============================================================
st.header("6. Рекомендации")
for rec in recommendations_for(row, date_cols):
    st.markdown(
        f"<div style='background:#fff7ed;border-left:4px solid #f97316;"
        f"padding:8px 12px;border-radius:8px;margin:4px 0'>{rec}</div>",
        unsafe_allow_html=True,
    )

# ============================================================
# ЭКСПОРТ
# ============================================================
st.header("7. Экспорт и печать")


def build_excel():
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        # Лист 1 — все оценки
        df.to_excel(writer, sheet_name="Все оценки", index=False)

        # Лист 2 — сводка (уже содержит «Білім сапасы %» и «Үлгерім %»)
        summary_df.to_excel(writer, sheet_name="Сводка", index=False)

        # Лист 3 — Топ
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

        # Лист 4 — рекомендации
        rec_rows = []
        for _, r in df.iterrows():
            for rec in recommendations_for(r, date_cols):
                rec_rows.append({"Ученик": r["Ученик"], "Рекомендация": rec})
        if not rec_rows:
            rec_rows = [{"Ученик": "—", "Рекомендация": "Нет данных"}]
        pd.DataFrame(rec_rows).to_excel(writer, sheet_name="Рекомендации", index=False)

        # Форматирование
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
