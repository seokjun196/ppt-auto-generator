import streamlit as st
import pandas as pd
import copy
import io
import re
import lxml.etree as etree
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ── 페이지 설정 ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PPT 자동 생성기",
    page_icon="📊",
    layout="centered",
)

st.title("📊 PPT 자동 생성기")
st.caption("엑셀 데이터를 PPT 양식에 자동으로 입력합니다.")

with st.expander("📖 사용 방법", expanded=False):
    st.markdown("""
    **엑셀 파일 형식**

    | 직책 | 직책 폰트 | 직책 글씨크기 | 이름 | 이름 폰트 | 이름 글씨크기 |
    |------|----------|------------|------|----------|------------|
    | 대표이사 | 맑은 고딕 | 24 | 홍길동 | 맑은 고딕 | 36 |
    | 부장 | Arial | 20 | 김철수 | Arial | 28 |

    - 1행은 **헤더(컬럼명)** 입니다.
    - 2행부터 데이터가 시작되며, 각 행이 PPT 슬라이드 한 장에 해당합니다.

    **PPT 양식 파일**
    - 글상자 안에 `{직책}`, `{이름}` 플레이스홀더가 있어야 합니다.
    - 이미지는 모든 슬라이드에 동일하게 반복됩니다.
    - 엑셀 행 수만큼 슬라이드가 생성됩니다.
    """)

# ── 파일 업로드 ──────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    st.subheader("1️⃣ 엑셀 파일")
    excel_file = st.file_uploader(
        "엑셀 파일을 업로드하세요", type=["xlsx", "xls"], key="excel",
        help="직책, 직책 폰트, 직책 글씨크기, 이름, 이름 폰트, 이름 글씨크기 컬럼이 필요합니다."
    )
with col2:
    st.subheader("2️⃣ PPT 양식 파일")
    ppt_file = st.file_uploader(
        "PPT 파일을 업로드하세요", type=["pptx"], key="ppt",
        help="{직책}, {이름} 플레이스홀더가 글상자에 포함되어 있어야 합니다."
    )

# ── 엑셀 미리보기 ────────────────────────────────────────────────────────────
df = None
if excel_file:
    try:
        df = pd.read_excel(excel_file)
        df.columns = df.columns.str.strip()
        required_cols = ["직책", "직책 폰트", "직책 글씨크기", "이름", "이름 폰트", "이름 글씨크기"]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            st.error(f"❌ 엑셀 파일에 다음 컬럼이 없습니다: {', '.join(missing)}")
            df = None
        else:
            st.success(f"✅ 엑셀 파일 로드 완료 — {len(df)}개 행 (슬라이드 {len(df)}장 생성 예정)")
            st.dataframe(df, use_container_width=True)
    except Exception as e:
        st.error(f"❌ 엑셀 파일을 읽는 중 오류 발생: {e}")

# ── 상수 ─────────────────────────────────────────────────────────────────────
PLACEHOLDER_PATTERN = re.compile(r"\{(직책|이름)\}")
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

# ── 헬퍼 함수 ─────────────────────────────────────────────────────────────────

def get_font_name(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    return str(val).strip()

def get_font_size(val):
    try:
        return Pt(float(val))
    except Exception:
        return None

def _set_run_text_xml(run, text):
    """<a:t> 텍스트만 교체 — <a:rPr> 건드리지 않음"""
    t_elem = run._r.find(f"{{{_A}}}t")
    if t_elem is not None:
        t_elem.text = text
    else:
        t_elem = etree.SubElement(run._r, f"{{{_A}}}t")
        t_elem.text = text

def _set_run_font_xml(run, font_name, font_size):
    """<a:rPr> 직접 수정으로 폰트명·크기 적용"""
    rPr = run._r.find(f"{{{_A}}}rPr")
    if rPr is None:
        rPr = etree.Element(f"{{{_A}}}rPr")
        run._r.insert(0, rPr)

    if font_name:
        latin = rPr.find(f"{{{_A}}}latin")
        if latin is None:
            latin = etree.SubElement(rPr, f"{{{_A}}}latin")
        latin.set("typeface", font_name)

    if font_size:
        rPr.set("sz", str(int(round(font_size.pt * 100))))

def process_paragraph(para, mapping):
    """
    paragraph 내 run을 합산 → 플레이스홀더 치환 → 폰트 적용.
    run이 몇 개로 쪼개져 있어도 첫 번째 run으로 병합 후 처리.
    """
    if not para.runs:
        return

    combined = "".join(r.text for r in para.runs)
    if not PLACEHOLDER_PATTERN.search(combined):
        return

    found_keys = list(dict.fromkeys(PLACEHOLDER_PATTERN.findall(combined)))

    # 첫 번째 run에 합산, 나머지 비움
    first_run = para.runs[0]
    _set_run_text_xml(first_run, combined)
    for r in para.runs[1:]:
        _set_run_text_xml(r, "")

    # 텍스트 치환
    new_text = combined
    for key in found_keys:
        info = mapping.get(key)
        if info:
            new_text = new_text.replace(f"{{{key}}}", info["value"])
    _set_run_text_xml(first_run, new_text)

    # 폰트 적용
    for key in found_keys:
        info = mapping.get(key)
        if info:
            _set_run_font_xml(first_run, info.get("font_name"), info.get("font_size"))

def iter_all_shapes(shapes):
    """
    그룹(GROUP) shape을 재귀적으로 내려가며 모든 shape을 순회.
    일반 slide.shapes 는 그룹 내부를 방문하지 않으므로 이 함수가 필수.
    """
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_all_shapes(shape.shapes)

def fill_slide(slide, mapping):
    """
    슬라이드의 모든 shape(그룹 내부 포함)을 순회하며
    {직책}/{이름} 플레이스홀더를 치환하고 폰트를 적용한다.
    """
    for shape in iter_all_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue

        # shape 전체 텍스트에 플레이스홀더가 있는지 먼저 확인
        full_text = "".join(
            r.text
            for para in shape.text_frame.paragraphs
            for r in para.runs
        )
        if not PLACEHOLDER_PATTERN.search(full_text):
            continue

        # paragraph마다 처리
        for para in shape.text_frame.paragraphs:
            process_paragraph(para, mapping)

# ── PPT 생성 ─────────────────────────────────────────────────────────────────

def generate_ppt(prs_template: Presentation, df: pd.DataFrame) -> bytes:
    prs_out = Presentation()
    prs_out.slide_width = prs_template.slide_width
    prs_out.slide_height = prs_template.slide_height

    template_slide = prs_template.slides[0]

    for _, row in df.iterrows():
        mapping = {
            "직책": {
                "value": str(row["직책"]).strip() if pd.notna(row["직책"]) else "",
                "font_name": get_font_name(row.get("직책 폰트")),
                "font_size": get_font_size(row.get("직책 글씨크기")),
            },
            "이름": {
                "value": str(row["이름"]).strip() if pd.notna(row["이름"]) else "",
                "font_name": get_font_name(row.get("이름 폰트")),
                "font_size": get_font_size(row.get("이름 글씨크기")),
            },
        }

        blank_layout = prs_out.slide_layouts[6]
        new_slide = prs_out.slides.add_slide(blank_layout)

        # spTree 전체를 한 번에 deepcopy → 슬라이드 간 XML 완전 격리
        cloned_sp_tree = copy.deepcopy(template_slide._element.spTree)
        dst_sp_tree = new_slide._element.spTree
        for child in list(dst_sp_tree):
            dst_sp_tree.remove(child)
        for child in list(cloned_sp_tree):
            dst_sp_tree.append(child)

        # 이미지 관계 복사
        for rel in template_slide.part.rels.values():
            if "image" in rel.reltype:
                new_slide.part.relate_to(rel.target_part, rel.reltype)

        fill_slide(new_slide, mapping)

    buf = io.BytesIO()
    prs_out.save(buf)
    buf.seek(0)
    return buf.read()

# ── 생성 버튼 ────────────────────────────────────────────────────────────────
st.divider()

if df is not None and ppt_file is not None:
    if st.button("🚀 PPT 생성하기", type="primary", use_container_width=True):
        with st.spinner("PPT 생성 중..."):
            try:
                prs_template = Presentation(io.BytesIO(ppt_file.read()))
                result_bytes = generate_ppt(prs_template, df)
                st.success(f"✅ PPT 생성 완료! 슬라이드 {len(df)}장이 만들어졌습니다.")
                st.download_button(
                    label="⬇️ PPT 다운로드",
                    data=result_bytes,
                    file_name="output.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"❌ PPT 생성 중 오류 발생: {e}")
                st.exception(e)
else:
    st.info("📂 엑셀 파일과 PPT 양식 파일을 모두 업로드하면 생성 버튼이 활성화됩니다.")
