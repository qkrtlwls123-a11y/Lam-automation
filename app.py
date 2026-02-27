import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData

TEMPLATE_PATH = Path(__file__).with_name("template.pptx")


def to_100_scale(avg_4_scale: float) -> float:
    return (avg_4_scale / 4.0) * 100.0


def detect_question_columns(df: pd.DataFrame) -> list[str]:
    headers = list(df.columns)
    picked: list[str] = []

    def normalize(text: str) -> str:
        return re.sub(r"\s+", "", str(text).strip().lower())

    normalized = {h: normalize(h) for h in headers}

    for i in range(1, 11):
        candidates = []
        for h in headers:
            n = normalized[h]
            if n in {str(i), f"q{i}", f"문항{i}", f"질문{i}"}:
                candidates.append(h)
            elif re.search(rf"(^|\D){i}(\D|$)", n) and ("q" in n or "문항" in n or "질문" in n):
                candidates.append(h)
        if candidates:
            picked.append(candidates[0])

    if len(picked) == 10:
        return picked

    numeric_like = []
    for h in headers:
        series = pd.to_numeric(df[h], errors="coerce").dropna()
        if series.empty:
            continue
        if series.between(1, 4).mean() >= 0.8:
            numeric_like.append(h)

    if len(numeric_like) < 10:
        raise ValueError("문항 데이터 열을 10개 이상 찾지 못했습니다. Q1~Q10 열명을 확인해주세요.")

    return numeric_like[:10]


def preprocess(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str], dict[int, dict[str, float]]]:
    q_cols = detect_question_columns(df)
    q_df = df[q_cols].copy()
    q_df.columns = [f"Q{i}" for i in range(1, 11)]
    q_df = q_df.apply(pd.to_numeric, errors="coerce")

    valid = q_df.dropna(how="all")
    if valid.empty:
        raise ValueError("유효한 응답 데이터가 없습니다.")

    sub_4 = valid.mean(axis=0)
    sub_100 = sub_4.apply(to_100_scale)

    total_avg_01 = to_100_scale(valid[[f"Q{i}" for i in range(1, 6)]].mean(axis=1).mean())
    total_avg_02 = to_100_scale(valid[["Q6", "Q7", "Q8"]].mean(axis=1).mean())
    total_avg_03 = to_100_scale(valid[["Q9", "Q10"]].mean(axis=1).mean())
    total_avg_04 = to_100_scale(valid[["Q2", "Q3", "Q4", "Q5"]].mean(axis=1).mean())
    total_avg_00 = to_100_scale(valid.mean(axis=1).mean())

    placeholders: dict[str, str] = {
        "total_avg_00": f"{total_avg_00:.1f}",
        "total_avg_01": f"{total_avg_01:.1f}",
        "total_avg_02": f"{total_avg_02:.1f}",
        "total_avg_03": f"{total_avg_03:.1f}",
        "total_avg_04": f"{total_avg_04:.1f}",
    }

    for i in range(1, 11):
        placeholders[f"sub_avg_{i:02d}"] = f"{sub_100[f'Q{i}']:.1f}"

    distributions: dict[int, dict[str, float]] = {}
    total_n = len(valid)
    for i in range(1, 11):
        s = valid[f"Q{i}"]
        counts = s.value_counts(dropna=False)
        c1 = int(counts.get(1.0, 0))
        c2 = int(counts.get(2.0, 0))
        c3 = int(counts.get(3.0, 0))
        c4 = int(counts.get(4.0, 0))
        distributions[i] = {
            "count_1": c1,
            "count_2": c2,
            "count_3": c3,
            "count_4": c4,
            "pct_1": (c1 / total_n * 100) if total_n else 0,
            "pct_2": (c2 / total_n * 100) if total_n else 0,
            "pct_3": (c3 / total_n * 100) if total_n else 0,
            "pct_4": (c4 / total_n * 100) if total_n else 0,
            "total": total_n,
        }

    return valid, placeholders, distributions


def replace_chart_data(shape, categories: list[str], values: list[float]) -> None:
    if not getattr(shape, "has_chart", False):
        return
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("", values)
    shape.chart.replace_data(chart_data)


def update_chart_0(shape, placeholders: dict[str, str]) -> None:
    values = [
        float(placeholders["total_avg_00"]),
        float(placeholders["total_avg_01"]),
        float(placeholders["sub_avg_01"]),
        float(placeholders["sub_avg_02"]),
        float(placeholders["sub_avg_03"]),
        float(placeholders["sub_avg_04"]),
        float(placeholders["sub_avg_05"]),
        float(placeholders["sub_avg_06"]),
        float(placeholders["sub_avg_07"]),
        float(placeholders["sub_avg_08"]),
        float(placeholders["sub_avg_09"]),
        float(placeholders["sub_avg_10"]),
        float(placeholders["total_avg_02"]),
        float(placeholders["total_avg_03"]),
    ]

    default_categories = [
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
    replace_chart_data(shape, default_categories, values)


def update_question_chart(shape, dist: dict[str, float]) -> None:
    categories = ["1점", "2점", "3점", "4점"]
    values = [dist["pct_1"], dist["pct_2"], dist["pct_3"], dist["pct_4"]]
    replace_chart_data(shape, categories, values)


def update_table(shape, dist: dict[str, float]) -> None:
    if not shape.has_table:
        return

    table = shape.table
    rows_data = [
        (dist["count_4"], dist["pct_4"]),
        (dist["count_3"], dist["pct_3"]),
        (dist["count_2"], dist["pct_2"]),
        (dist["count_1"], dist["pct_1"]),
    ]

    for idx, (count, pct) in enumerate(rows_data):
        r, c = (1 + idx), 1
        if r < len(table.rows) and c < len(table.columns):
            table.cell(r, c).text = f"{count}명({pct:.0f}%)"

    if 5 < len(table.rows) and 1 < len(table.columns):
        total = dist["total"]
        table.cell(5, 1).text = f"{total}명(100%)"


def replace_placeholders_in_shape(shape, placeholder_values: dict[str, str]) -> None:
    if shape.has_text_frame:
        text = shape.text_frame.text
        for key, value in placeholder_values.items():
            text = text.replace(f"{{{{{key}}}}}", value)
        shape.text_frame.text = text

    if hasattr(shape, "shapes"):
        for sub_shape in shape.shapes:
            replace_placeholders_in_shape(sub_shape, placeholder_values)


def build_ppt(excel_bytes: bytes, class_name: str) -> bytes:
    df = pd.read_excel(io.BytesIO(excel_bytes))
    _, placeholders, distributions = preprocess(df)
    placeholders["class_name"] = class_name

    prs = Presentation(str(TEMPLATE_PATH))

    for slide in prs.slides:
        for shape in slide.shapes:
            name = (shape.name or "").strip().lower()

            if name in {"차트 0", "chart 0"}:
                update_chart_0(shape, placeholders)
            else:
                m_chart = re.search(r"(차트|chart)\s*(\d+)", name)
                if m_chart:
                    idx = int(m_chart.group(2))
                    if 1 <= idx <= 10:
                        update_question_chart(shape, distributions[idx])

            m_table = re.search(r"(표|table)\s*(\d+)", name)
            if m_table:
                idx = int(m_table.group(2))
                if 1 <= idx <= 10:
                    update_table(shape, distributions[idx])

            replace_placeholders_in_shape(shape, placeholders)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()


def main() -> None:
    st.set_page_config(page_title="설문 PPT 자동 생성기", layout="wide")
    st.title("설문 결과 PPT 자동 생성기")
    st.write("엑셀(raw data)과 과정명을 입력하면 PPT를 즉시 생성해 다운로드할 수 있습니다.")

    class_name = st.text_input("과정명", placeholder="예: 2026 리더십 과정")
    uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

    if st.button("PPT 생성", type="primary"):
        if not class_name.strip():
            st.error("과정명을 입력해주세요.")
            return
        if uploaded_file is None:
            st.error("엑셀 파일을 업로드해주세요.")
            return
        if not TEMPLATE_PATH.exists():
            st.error("template.pptx 파일을 찾지 못했습니다.")
            return

        try:
            ppt_bytes = build_ppt(uploaded_file.getvalue(), class_name.strip())
        except Exception as exc:
            st.exception(exc)
            return

        st.success("PPT 생성이 완료되었습니다.")
        st.download_button(
            label="결과 PPT 다운로드",
            data=ppt_bytes,
            file_name=f"{class_name.strip()}_설문결과.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )


if __name__ == "__main__":
    main()
