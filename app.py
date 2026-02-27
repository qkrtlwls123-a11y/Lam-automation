import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData

TEMPLATE_PATH = "template.pptx"


def _find_question_columns(df: pd.DataFrame) -> List[str]:
    q_regex = re.compile(r"^\s*q\s*([1-9]|10)\s*$", re.IGNORECASE)
    q_cols: List[Tuple[int, str]] = []

    for col in df.columns:
        match = q_regex.match(str(col))
        if match:
            q_cols.append((int(match.group(1)), col))

    if len(q_cols) >= 10:
        q_cols = sorted(q_cols, key=lambda x: x[0])
        return [col for _, col in q_cols[:10]]

    numeric_like = []
    for col in df.columns:
        series = pd.to_numeric(df[col], errors="coerce")
        valid = series.dropna()
        if valid.empty:
            continue
        if valid.between(1, 4).mean() >= 0.9:
            numeric_like.append(col)

    if len(numeric_like) < 10:
        raise ValueError("Q1~Q10 문항 컬럼을 찾지 못했습니다. 엑셀 컬럼명을 확인해 주세요.")

    return numeric_like[:10]


def _score_100(avg_4_scale: float) -> float:
    return (avg_4_scale / 4.0) * 100.0


def _format_pct(value: float) -> str:
    return f"{value:.1f}%"


def _format_score(value: float) -> str:
    return f"{value:.1f}"


def preprocess(df: pd.DataFrame) -> Tuple[Dict[str, float], Dict[int, Dict[int, Tuple[int, float]]]]:
    q_cols = _find_question_columns(df)
    sub_scores: Dict[int, float] = {}
    distributions: Dict[int, Dict[int, Tuple[int, float]]] = {}

    for idx, col in enumerate(q_cols, start=1):
        series = pd.to_numeric(df[col], errors="coerce")
        series = series[series.between(1, 4)]
        if series.empty:
            raise ValueError(f"{col} 컬럼에서 1~4점 응답을 찾지 못했습니다.")

        sub_scores[idx] = _score_100(series.mean())
        counts = series.value_counts().to_dict()
        total = int(series.count())
        distributions[idx] = {}
        for score in [1, 2, 3, 4]:
            cnt = int(counts.get(score, 0))
            pct = (cnt / total * 100.0) if total else 0.0
            distributions[idx][score] = (cnt, pct)

    placeholder_values: Dict[str, float] = {}
    for i in range(1, 11):
        placeholder_values[f"sub_avg_{i:02d}"] = sub_scores[i]

    placeholder_values["total_avg_00"] = sum(sub_scores.values()) / 10
    placeholder_values["total_avg_01"] = sum(sub_scores[i] for i in range(1, 6)) / 5
    placeholder_values["total_avg_02"] = sum(sub_scores[i] for i in range(6, 9)) / 3
    placeholder_values["total_avg_03"] = sum(sub_scores[i] for i in range(9, 11)) / 2
    placeholder_values["total_avg_04"] = sum(sub_scores[i] for i in range(2, 6)) / 4

    return placeholder_values, distributions


def _extract_numbered_name(shape_name: str, patterns: List[str]) -> int:
    for pattern in patterns:
        match = re.search(pattern, shape_name, flags=re.IGNORECASE)
        if match:
            return int(match.group(1))
    return -1


def update_chart_0(shape, values: Dict[str, float]) -> None:
    chart_data = CategoryChartData()
    categories = [
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
    data_values = [
        values["total_avg_00"],
        values["total_avg_01"],
        values["sub_avg_01"],
        values["sub_avg_02"],
        values["sub_avg_03"],
        values["sub_avg_04"],
        values["sub_avg_05"],
        values["sub_avg_06"],
        values["sub_avg_07"],
        values["sub_avg_08"],
        values["sub_avg_09"],
        values["sub_avg_10"],
        values["total_avg_02"],
        values["total_avg_03"],
    ]
    chart_data.categories = categories
    chart_data.add_series("점수", data_values)
    shape.chart.replace_data(chart_data)


def update_chart_question(shape, question_no: int, distributions: Dict[int, Dict[int, Tuple[int, float]]]) -> None:
    dist = distributions[question_no]
    chart_data = CategoryChartData()
    chart_data.categories = ["1점", "2점", "3점", "4점"]
    chart_data.add_series("응답 비율", [dist[1][1], dist[2][1], dist[3][1], dist[4][1]])
    shape.chart.replace_data(chart_data)


def update_table_question(shape, question_no: int, distributions: Dict[int, Dict[int, Tuple[int, float]]]) -> None:
    table = shape.table
    dist = distributions[question_no]
    total = sum(dist[score][0] for score in [1, 2, 3, 4])

    rows = [
        (2, 2, 4),
        (3, 2, 3),
        (4, 2, 2),
        (5, 2, 1),
    ]

    for row_idx, col_idx, score in rows:
        count, pct = dist[score]
        table.cell(row_idx - 1, col_idx - 1).text = f"{count}명({pct:.0f}%)"

    table.cell(6 - 1, 2 - 1).text = f"{total}명"


def replace_placeholders_in_shape(shape, replacements: Dict[str, str]) -> None:
    if not shape.has_text_frame:
        return

    text = shape.text_frame.text
    for key, value in replacements.items():
        text = text.replace(f"{{{{{key}}}}}", value)

    shape.text_frame.clear()
    shape.text_frame.paragraphs[0].text = text


def generate_ppt(excel_file, class_name: str) -> bytes:
    df = pd.read_excel(excel_file)
    values, distributions = preprocess(df)

    prs = Presentation(TEMPLATE_PATH)

    replacements = {k: _format_score(v) for k, v in values.items()}
    replacements["class_name"] = class_name

    for slide in prs.slides:
        for shape in slide.shapes:
            replace_placeholders_in_shape(shape, replacements)

            if shape.has_chart:
                chart_no = _extract_numbered_name(
                    shape.name,
                    [r"(?:차트|chart)\s*(\d+)", r"(?:차트|chart)[^0-9]*(\d+)"],
                )
                if chart_no == 0:
                    update_chart_0(shape, values)
                elif 1 <= chart_no <= 10:
                    update_chart_question(shape, chart_no, distributions)

            if shape.has_table:
                table_no = _extract_numbered_name(
                    shape.name,
                    [r"(?:표|table)\s*(\d+)", r"(?:표|table)[^0-9]*(\d+)"],
                )
                if 1 <= table_no <= 10:
                    update_table_question(shape, table_no, distributions)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output.read()


def main() -> None:
    st.set_page_config(page_title="교육 만족도 PPT 생성기", layout="centered")
    st.title("교육 만족도 PPT 자동 생성")
    st.write("엑셀(raw data)과 과정명을 입력하면 템플릿 PPT를 채워 다운로드합니다.")

    uploaded_file = st.file_uploader("엑셀 파일(.xlsx)", type=["xlsx"])
    class_name = st.text_input("과정명 입력", placeholder="예: 2026 상반기 리더십 교육")

    if st.button("PPT 생성"):
        if not uploaded_file:
            st.error("엑셀 파일을 업로드해 주세요.")
            return
        if not class_name.strip():
            st.error("과정명을 입력해 주세요.")
            return

        try:
            ppt_bytes = generate_ppt(uploaded_file, class_name.strip())
            st.success("PPT 생성이 완료되었습니다.")
            st.download_button(
                label="결과 PPT 다운로드",
                data=ppt_bytes,
                file_name=f"{class_name.strip()}_결과보고서.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        except Exception as exc:
            st.exception(exc)


if __name__ == "__main__":
    main()
