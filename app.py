import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData

SCALE_MAX = 4
QUESTION_COUNT = 10


def to_percent_100(avg: float) -> float:
    return (avg / SCALE_MAX) * 100


def find_question_columns(df: pd.DataFrame) -> List[str]:
    q_regex = re.compile(r"^q\s*([1-9]|10)$", re.IGNORECASE)
    named = [c for c in df.columns if q_regex.match(str(c).strip())]
    if len(named) >= QUESTION_COUNT:
        named_sorted = sorted(named, key=lambda x: int(re.findall(r"\d+", str(x))[0]))
        return named_sorted[:QUESTION_COUNT]

    numeric_like = []
    for c in df.columns:
        series = pd.to_numeric(df[c], errors="coerce")
        valid = series.dropna()
        if not valid.empty and valid.between(1, 4).all():
            numeric_like.append(c)

    if len(numeric_like) < QUESTION_COUNT:
        raise ValueError("문항 데이터(Q1~Q10)를 찾지 못했습니다. Q1~Q10 컬럼 또는 1~4점 값 컬럼 10개가 필요합니다.")

    return numeric_like[:QUESTION_COUNT]


def preprocess(df: pd.DataFrame, q_cols: List[str]) -> Dict[str, float]:
    answers = df[q_cols].apply(pd.to_numeric, errors="coerce")
    answers = answers.where(answers.between(1, 4))

    sub_avgs = [answers[c].mean(skipna=True) for c in q_cols]
    total_avg_00 = to_percent_100(pd.Series(sub_avgs).mean(skipna=True))
    total_avg_01 = to_percent_100(pd.Series(sub_avgs[:5]).mean(skipna=True))
    total_avg_02 = to_percent_100(pd.Series(sub_avgs[5:8]).mean(skipna=True))
    total_avg_03 = to_percent_100(pd.Series(sub_avgs[8:10]).mean(skipna=True))
    total_avg_04 = to_percent_100(pd.Series(sub_avgs[1:5]).mean(skipna=True))

    result: Dict[str, float] = {
        "total_avg_00": total_avg_00,
        "total_avg_01": total_avg_01,
        "total_avg_02": total_avg_02,
        "total_avg_03": total_avg_03,
        "total_avg_04": total_avg_04,
    }
    for i, avg in enumerate(sub_avgs, start=1):
        result[f"sub_avg_{i:02d}"] = to_percent_100(avg)

    return result


def get_distribution(series: pd.Series) -> Tuple[Dict[int, int], Dict[int, float], int]:
    clean = pd.to_numeric(series, errors="coerce")
    clean = clean[clean.between(1, 4)]
    total = int(clean.shape[0])
    counts = {score: int((clean == score).sum()) for score in [1, 2, 3, 4]}
    if total == 0:
        rates = {score: 0.0 for score in [1, 2, 3, 4]}
    else:
        rates = {score: (counts[score] / total) * 100 for score in [1, 2, 3, 4]}
    return counts, rates, total


def replace_placeholders(prs: Presentation, values: Dict[str, str]) -> None:
    pattern = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}")

    def repl(txt: str) -> str:
        def _sub(m):
            key = m.group(1)
            return values.get(key, m.group(0))

        return pattern.sub(_sub, txt)

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        run.text = repl(run.text)

            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        cell.text = repl(cell.text)


def update_chart(shape, categories: List[str], values: List[float]) -> None:
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("", values)
    shape.chart.replace_data(chart_data)


def build_ppt(ppt_bytes: bytes, df: pd.DataFrame, class_name: str) -> bytes:
    q_cols = find_question_columns(df)
    metrics = preprocess(df, q_cols)

    question_dists = {}
    for i, col in enumerate(q_cols, start=1):
        question_dists[i] = get_distribution(df[col])

    prs = Presentation(io.BytesIO(ppt_bytes))

    chart0_categories = [
        "전체 평균",
        "과정 만족도",
        "문항1",
        "문항2",
        "문항3",
        "문항4",
        "문항5",
        "문항6",
        "문항7",
        "문항8",
        "문항9",
        "문항10",
        "강사 만족도",
        "운영 만족도",
    ]
    chart0_values = [
        metrics["total_avg_00"],
        metrics["total_avg_01"],
        *[metrics[f"sub_avg_{i:02d}"] for i in range(1, 11)],
        metrics["total_avg_02"],
        metrics["total_avg_03"],
    ]

    for slide in prs.slides:
        for shape in slide.shapes:
            name = getattr(shape, "name", "") or ""
            lname = name.lower().strip()

            if shape.has_chart:
                if lname in {"차트 0", "chart 0"}:
                    update_chart(shape, chart0_categories, chart0_values)
                else:
                    chart_match = re.match(r"^(차트|chart)\s*(\d+)$", lname)
                    if chart_match:
                        idx = int(chart_match.group(2))
                        if 1 <= idx <= QUESTION_COUNT:
                            _, rates, _ = question_dists[idx]
                            categories = ["1점", "2점", "3점", "4점"]
                            values = [rates[1], rates[2], rates[3], rates[4]]
                            update_chart(shape, categories, values)

            if shape.has_table:
                table_match = re.match(r"^(표|table)\s*(\d+)$", lname)
                if table_match:
                    idx = int(table_match.group(2))
                    if 1 <= idx <= QUESTION_COUNT:
                        counts, rates, total = question_dists[idx]
                        rows = [
                            (4, counts[4], rates[4]),
                            (3, counts[3], rates[3]),
                            (2, counts[2], rates[2]),
                            (1, counts[1], rates[1]),
                        ]
                        for row_offset, (_, cnt, rt) in enumerate(rows, start=2):
                            table_cell = shape.table.cell(row_offset - 1, 1)
                            table_cell.text = f"{cnt}명({rt:.0f}%)"
                        shape.table.cell(5, 1).text = f"{total}명(100%)" if total else "0명(0%)"

    text_values = {k: f"{v:.1f}" for k, v in metrics.items()}
    text_values["class_name"] = class_name
    replace_placeholders(prs, text_values)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.getvalue()


st.set_page_config(page_title="교육 만족도 PPT 생성기", layout="centered")
st.title("교육 만족도 PPT 자동 생성")
st.caption("엑셀 Raw Data(.xlsx)와 과정명을 입력하면, 데이터가 반영된 PPT를 즉시 다운로드할 수 있습니다.")

uploaded_excel = st.file_uploader("엑셀 Raw Data 업로드 (.xlsx)", type=["xlsx"])
uploaded_ppt = st.file_uploader("PPT 템플릿 업로드 (.pptx)", type=["pptx"])
class_name = st.text_input("과정명 입력", placeholder="예: 2026 상반기 리더십 과정")

if st.button("PPT 생성"):
    if not uploaded_excel or not uploaded_ppt or not class_name.strip():
        st.error("엑셀, PPT 템플릿, 과정명을 모두 입력해주세요.")
    else:
        try:
            df = pd.read_excel(uploaded_excel)
            generated = build_ppt(uploaded_ppt.getvalue(), df, class_name.strip())
            st.success("PPT 생성 완료")
            st.download_button(
                label="결과 PPT 다운로드",
                data=generated,
                file_name=f"{class_name.strip()}_결과보고서.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        except Exception as e:
            st.exception(e)
