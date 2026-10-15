# ============================================================
# ЧТЕНИЕ ФИО УЧИТЕЛЯ (лист 2, строка 1)
# ============================================================
def read_teacher_name(xls):
    """
    Читает ФИО учителя с листа № 2 (индекс 1) из первой строки.
    Ищет первую непустую ячейку, похожую на ФИО (2–3 слова с заглавной буквы).
    """
    info = []
    if len(xls.sheet_names) < 2:
        info.append("Листов меньше 2 — ФИО учителя не найдено.")
        return "", info

    second_sheet = xls.sheet_names[1]  # лист № 2
    try:
        raw = pd.read_excel(xls, sheet_name=second_sheet, header=None, nrows=3)
    except Exception as e:
        info.append(f"Не удалось прочитать лист «{second_sheet}»: {e}")
        return "", info

    if raw.empty:
        info.append(f"Лист «{second_sheet}» пуст.")
        return "", info

    # Берём первую строку
    first_row = raw.iloc[0].tolist()
    for cell in first_row:
        s = str(cell).strip()
        if not s or s.lower() in ("nan", "none"):
            continue
        # 2–3 слова с заглавной буквы, кириллица
        if re.match(r"^([А-ЯЁӘҒҚҢӨҰҮҺІ][а-яёәғқңөұүһі\-]+)(\s+[А-ЯЁӘҒҚҢӨҰҮҺІ][а-яёәғқңөұүһі\-]+){1,3}$", s):
            info.append(f"ФИО учителя найдено на листе «{second_sheet}»: {s}")
            return s, info

    info.append(f"ФИО учителя не найдено в первой строке листа «{second_sheet}».")
    return "", info
