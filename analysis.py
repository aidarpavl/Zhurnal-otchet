"""
Анализ данных журнала: сильные/слабые темы, рекомендации, вопросы,
топ лучших и топ худших учеников.
"""
from typing import List, Dict

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


def _to_score(mark: str):
    """'4' → 4; 'А'/'С'/'Н' → None (пропуск)."""
    if mark is None:
        return None
    m = str(mark).strip().upper()
    if m in ("А", "С", "Н"):
        return None
    try:
        return int(m)
    except ValueError:
        return None


def build_dataframe(students_raw: List[Dict], topics=TOPICS):
    """
    Преобразует распознанные строки в структуру {name: {topic: value/miss}}.
    Возвращает список словарей — по одному на ученика.
    """
    result = []
    for s in students_raw:
        row = {"name": s["name"]}
        marks = s.get("marks", [])
        for i, topic in enumerate(topics):
            mark = marks[i] if i < len(marks) else ""
            mark = str(mark).strip().upper()
            if mark in ("А", "С", "Н"):
                row[topic] = "А" if mark == "А" else ("С" if mark == "С" else "")
            else:
                val = _to_score(mark)
                row[topic] = val if val is not None else ""
        result.append(row)
    return result


def analyze_student(row: Dict, topics=TOPICS) -> Dict:
    """Возвращает средний балл, сильные и слабые темы, пропуски."""
    scores = []
    strong, weak, missing = [], [], []
    for t in topics:
        v = row.get(t, "")
        if v == "А" or v == "С":
            missing.append({"topic": t, "reason": v})
        elif isinstance(v, int):
            scores.append(v)
            if v >= 4:
                strong.append({"topic": t, "value": v})
            elif v <= 3:
                weak.append({"topic": t, "value": v})
    avg = round(sum(scores) / len(scores), 2) if scores else 0.0
    return {
        "avg": avg,
        "strong": strong,
        "weak": weak,
        "missing": missing,
        "scores": scores,
    }


def top_best(rows: List[Dict], n: int = 3, topics=TOPICS) -> List[Dict]:
    scored = [(r["name"], analyze_student(r, topics)["avg"]) for r in rows]
    scored.sort(key=lambda x: x[1], reverse=True)
    return scored[:n]


def top_worst(rows: List[Dict], n: int = 3, topics=TOPICS) -> List[Dict]:
    scored = [(r["name"], analyze_student(r, topics)["avg"]) for r in rows]
    scored = [(n_, a) for n_, a in scored if a > 0]
    scored.sort(key=lambda x: x[1])
    return scored[:n]


def recommendations(row: Dict, topics=TOPICS) -> List[str]:
    """Текстовые рекомендации по слабым темам и пропускам."""
    a = analyze_student(row, topics)
    recs = []
    for w in a["weak"]:
        recs.append(
            f"📉 «{w['topic']}» — балл {w['value']}. "
            f"Повторить теорию, разобрать типичные задачи, решить 5–7 упражнений."
        )
    for m in a["missing"]:
        reason = "болел" if m["reason"] == "А" else "уважительная причина"
        recs.append(
            f"🩺 «{m['topic']}» — пропуск ({reason}). "
            f"Изучить тему самостоятельно и ответить на контрольные вопросы ниже."
        )
    if not recs:
        recs.append("✅ Слабых тем и пропусков нет. Отличная работа!")
    return recs


def questions_for_topics(topics_list: List[str]) -> Dict[str, List[str]]:
    """Возвращает по 5 вопросов на каждую тему."""
    return {t: QUESTION_BANK.get(t, []) for t in topics_list}