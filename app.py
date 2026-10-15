def analyze_student(row, date_cols):
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

    # Білім сапасы — доля оценок 7–10 (отлично/хорошо)
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
