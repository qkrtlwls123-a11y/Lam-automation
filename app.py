import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData

TEMPLATE_PATH = "template.pptx"


def identify_question_columns(df: pd.DataFrame, required: int = 10) -> List[str]:
    """문항(1~4점 척도) 컬럼 10개를 찾아 반환한다."""
    normalized = {str(col).strip().lower(): col for col in df.columns}
    selected: List[str] = []

    for idx in range(1, required + 1):
        candidates = [
            f"q{idx}",
            f"q{idx}.",
            f"q{idx} ",
            f"문항{idx}",
            f"문항 {idx}",
            f"{idx}번",
            f"{idx}",
        ]
        found = None
        for key, original in normalized.items():
            if key in candidates or re.fullmatch(rf"q\s*{idx}", key):
                found = original
                break
        if found and found not in selected:
            selected.append(found)

    if len(selected) >= required:
        return selected[:required]

    score_like_cols: List[str] = []
    for col in df.columns:
        series = pd.to_numeric(df[col], errors="coerce").dropna()
        if series.empty:
            continue
        if series.between(1, 4).all():
            score_like_cols.append(col)

    for col in score_like_cols:
        if col not in selected:
            selected.append(col)
        if len(selected) == required:
            break

    if len(selected) < required:
        raise ValueError(
            "1~4점 척도 문항 컬럼 10개를 찾지 못했습니다. 엑셀 컬럼명(Q1~Q10 등)을 확인해주세요."
        )

    return selected


def to_100_scale(raw_avg: float) -> float:
    return (raw_avg / 4.0) * 100.0


def format_score(value: float) -> str:
    return f"{value:.1f}"


def compute_metrics(df: pd.DataFrame, question_cols: List[str]) -> Tuple[Dict[str, float], Dict[int, Dict[int, float]], Dict[int, Dict[int, int]], int]:
    numeric = df[question_cols].apply(pd.to_numeric, errors="coerce")
    respondent_count = int(numeric.dropna(how="all").shape[0])

    sub_avgs_raw = [numeric[col].mean() for col in question_cols]
    sub_avgs_100 = [to_100_scale(v) if pd.notna(v) else 0.0 for v in sub_avgs_raw]

    total_avg_00 = to_100_scale(numeric.mean(axis=1).mean()) if not numeric.empty else 0.0
    total_avg_01 = to_100_scale(numeric.iloc[:, 0:5].mean(axis=1).mean())
    total_avg_02 = to_100_scale(numeric.iloc[:, 5:8].mean(axis=1).mean())
    total_avg_03 = to_100_scale(numeric.iloc[:, 8:10].mean(axis=1).mean())
    total_avg_04 = to_100_scale(numeric.iloc[:, 1:5].mean(axis=1).mean())

    placeholders: Dict[str, float] = {
        "total_avg_00": total_avg_00,
        "total_avg_01": total_avg_01,
        "total_avg_02": total_avg_02,
        "total_avg_03": total_avg_03,
        "total_avg_04": total_avg_04,
    }

    for idx, value in enumerate(sub_avgs_100, start=1):
        placeholders[f"sub_avg_{idx:02d}"] = value

    percentages_by_question: Dict[int, Dict[int, float]] = {}
    counts_by_question: Dict[int, Dict[int, int]] = {}

    for q_idx, col in enumerate(question_cols, start=1):
        answers = pd.to_numeric(numeric[col], errors="coerce").dropna().astype(int)
        total = len(answers)
        counts = {score: int((answers == score).sum()) for score in [1, 2, 3, 4]}
        percentages = {
            score: ((counts[score] / total) * 100.0 if total else 0.0)
            for score in [1, 2, 3, 4]
        }
        counts_by_question[q_idx] = counts
        percentages_by_question[q_idx] = percentages

    return placeholders, percentages_by_question, counts_by_question, respondent_count


def replace_text_placeholders(prs: Presentation, replacements: Dict[str, str]) -> None:
    pattern = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}")

    def _replace_text(text: str) -> str:
        def repl(match: re.Match) -> str:
            key = match.group(1)
            return replacements.get(key, match.group(0))

        return pattern.sub(repl, text)

    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame is not None:
                for paragraph in shape.text_frame.paragraphs:
                    if not paragraph.text:
                        continue
                    new_text = _replace_text(paragraph.text)
                    if new_text != paragraph.text:
                        paragraph.text = new_text

            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for paragraph in cell.text_frame.paragraphs:
                            if not paragraph.text:
                                continue
                            new_text = _replace_text(paragraph.text)
                            if new_text != paragraph.text:
                                paragraph.text = new_text


def update_chart_0(shape, placeholders: Dict[str, float]) -> None:
    chart_data = CategoryChartData()
    chart_data.categories = [
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
    chart_data.add_series(
        "점수",
        (
            placeholders["total_avg_00"],
            placeholders["total_avg_01"],
            placeholders["sub_avg_01"],
            placeholders["sub_avg_02"],
            placeholders["sub_avg_03"],
            placeholders["sub_avg_04"],
            placeholders["sub_avg_05"],
            placeholders["sub_avg_06"],
            placeholders["sub_avg_07"],
            placeholders["sub_avg_08"],
            placeholders["sub_avg_09"],
            placeholders["sub_avg_10"],
            placeholders["total_avg_02"],
            placeholders["total_avg_03"],
        ),
    )
    shape.chart.replace_data(chart_data)


def update_question_chart(shape, question_idx: int, percentages_by_question: Dict[int, Dict[int, float]]) -> None:
    percentages = percentages_by_question[question_idx]
    chart_data = CategoryChartData()
    chart_data.categories = ["1점", "2점", "3점", "4점"]
    chart_data.add_series("응답 비율(%)", (percentages[1], percentages[2], percentages[3], percentages[4]))
    shape.chart.replace_data(chart_data)


def update_question_table(shape, question_idx: int, counts_by_question: Dict[int, Dict[int, int]]) -> None:
    table = shape.table
    counts = counts_by_question[question_idx]
    total = sum(counts.values())

    def pct(score: int) -> int:
        return int(round((counts[score] / total) * 100)) if total else 0

    rows = [
        f"{counts[4]}명({pct(4)}%)",
        f"{counts[3]}명({pct(3)}%)",
        f"{counts[2]}명({pct(2)}%)",
        f"{counts[1]}명({pct(1)}%)",
        f"{total}명(100%)",
    ]

    for offset, value in enumerate(rows, start=1):
        if len(table.rows) > offset and len(table.columns) > 1:
            table.cell(offset, 1).text = value


def populate_ppt(
    excel_bytes: bytes,
    class_name: str,
) -> bytes:
    df = pd.read_excel(io.BytesIO(excel_bytes))
    question_cols = identify_question_columns(df)

    placeholders, percentages_by_question, counts_by_question, respondent_count = compute_metrics(df, question_cols)

    replacements: Dict[str, str] = {
        key: format_score(value) for key, value in placeholders.items()
    }
    replacements["class_name"] = class_name.strip() if class_name.strip() else "과정명 미입력"
    replacements["respondent_count"] = str(respondent_count)

    prs = Presentation(TEMPLATE_PATH)

    for slide in prs.slides:
        for shape in slide.shapes:
            name = (shape.name or "").strip().lower()

            if shape.has_chart:
                if name in {"차트 0", "chart 0"}:
                    update_chart_0(shape, placeholders)
                else:
                    chart_match = re.match(r"(?:차트|chart)\s*(\d+)$", name)
                    if chart_match:
                        idx = int(chart_match.group(1))
                        if 1 <= idx <= 10:
                            update_question_chart(shape, idx, percentages_by_question)

            if shape.has_table:
                table_match = re.match(r"(?:표|table)\s*(\d+)$", name)
                if table_match:
                    idx = int(table_match.group(1))
                    if 1 <= idx <= 10:
                        update_question_table(shape, idx, counts_by_question)

    replace_text_placeholders(prs, replacements)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()


def main() -> None:
    st.set_page_config(page_title="만족도 PPT 자동 생성기", layout="centered")
    st.title("만족도 조사 PPT 자동 생성기")
    st.write("엑셀(raw data) 업로드 후 버튼을 누르면 템플릿 기반 PPT를 즉시 다운로드할 수 있습니다.")

    class_name = st.text_input("과정명", placeholder="예: 2026년 신입사원 교육")
    uploaded_excel = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

    if st.button("PPT 생성", type="primary"):
        if not uploaded_excel:
            st.error("엑셀 파일을 먼저 업로드해주세요.")
            return

        try:
            ppt_bytes = populate_ppt(uploaded_excel.read(), class_name)
        except Exception as exc:  # noqa: BLE001
            st.exception(exc)
            return

        st.success("PPT가 생성되었습니다. 아래 버튼으로 바로 다운로드하세요.")
        st.download_button(
            label="결과 PPT 다운로드",
            data=ppt_bytes,
            file_name=f"{class_name.strip() or 'result'}_만족도_보고서.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )


if __name__ == "__main__":
    main()
