import streamlit as st
import pandas as pd
import copy
import io
import re
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

# ── 페이지 설정 ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PPT 자동 생성기",
    page_icon="📊",
    layout="centered",
)

st.title("📊 PPT 자동 생성기")
st.caption("엑셀 데이터를 PPT 양식에 자동으로 입력합니다.")

# ── 사용 방법 안내 ────────────────────────────────────────────────────────────
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
        "엑셀 파일을 업로드하세요",
        type=["xlsx", "xls"],
        key="excel",
        help="직책, 직책 폰트, 직책 글씨크기, 이름, 이름 폰트, 이름 글씨크기 컬럼이 필요합니다."
    )

with col2:
    st.subheader("2️⃣ PPT 양식 파일")
    ppt_file = st.file_uploader(
        "PPT 파일을 업로드하세요",
        type=["pptx"],
        key="ppt",
        help="{직책}, {이름} 플레이스홀더가 글상자에 포함되어 있어야 합니다."
    )

# ── 엑셀 미리보기 ────────────────────────────────────────────────────────────
df = None
if excel_file:
    try:
        df = pd.read_excel(excel_file)

        # 컬럼명 정규화 (공백 제거)
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

# ── PPT 생성 핵심 함수 ────────────────────────────────────────────────────────
PLACEHOLDER_PATTERN = re.compile(r"\{(직책|이름)\}")

def get_font_name(val):
    """None / NaN 처리"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    return str(val).strip()

def get_font_size(val):
    """None / NaN 처리 후 Pt로 변환"""
    try:
        return Pt(float(val))
    except Exception:
        return None

def replace_text_in_run(run, mapping):
    """run 안의 {직책}/{이름}을 치환하고 폰트를 적용."""
    text = run.text
    for key, info in mapping.items():
        placeholder = f"{{{key}}}"
        if placeholder in text:
            text = text.replace(placeholder, info["value"])
            # 폰트 적용
            if info.get("font_name"):
                run.font.name = info["font_name"]
            if info.get("font_size"):
                run.font.size = info["font_size"]
    run.text = text

def fill_slide(slide, mapping):
    """슬라이드 내 모든 텍스트프레임의 플레이스홀더를 치환."""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        tf = shape.text_frame
        # 슬라이드에 {직책} 또는 {이름} 이 있는지 먼저 확인
        full_text = "".join(run.text for para in tf.paragraphs for run in para.runs)
        if not PLACEHOLDER_PATTERN.search(full_text):
            continue

        for para in tf.paragraphs:
            # paragraph 레벨에서 합쳐진 텍스트 확인
            para_text = "".join(r.text for r in para.runs)
            if not PLACEHOLDER_PATTERN.search(para_text):
                continue

            # run이 분리된 경우 합쳐서 처리
            if len(para.runs) > 1:
                combined = "".join(r.text for r in para.runs)
                if PLACEHOLDER_PATTERN.search(combined):
                    # 첫 번째 run에 모아서 처리, 나머지 비움
                    first_run = para.runs[0]
                    first_run.text = combined
                    for r in para.runs[1:]:
                        r.text = ""
                    replace_text_in_run(first_run, mapping)
            else:
                for run in para.runs:
                    replace_text_in_run(run, mapping)

def generate_ppt(prs_template: Presentation, df: pd.DataFrame) -> bytes:
    """
    템플릿 첫 슬라이드를 기준으로 슬라이드를 복제해
    각 행의 데이터를 채워 넣는다.
    """
    # 새 Presentation 객체 (슬라이드 없이 시작)
    from pptx.util import Emu
    import lxml.etree as etree

    # 슬라이드 크기를 템플릿과 동일하게
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

        # 슬라이드 레이아웃(blank) 추가
        blank_layout = prs_out.slide_layouts[6]  # blank
        new_slide = prs_out.slides.add_slide(blank_layout)

        # 템플릿 슬라이드의 XML을 깊게 복사해 새 슬라이드에 붙여넣기
        template_xml = copy.deepcopy(template_slide._element)

        # spTree(shape tree) 복사
        new_sp_tree = new_slide._element.spTree
        template_sp_tree = template_xml.find(
            ".//{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"
            "/{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
            "spTree"
        )

        # namespace-aware 방법으로 spTree 가져오기
        PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
        DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

        cSld = template_xml.find(f"{{{PML_NS}}}cSld")
        if cSld is None:
            # 네임스페이스 없이 시도
            cSld = template_xml.find("cSld")

        # spTree를 직접 복사
        src_sp_tree = new_slide._element.spTree
        # 기존 요소 제거
        for child in list(src_sp_tree):
            src_sp_tree.remove(child)

        # 템플릿 슬라이드의 spTree 자식들 복사
        orig_sp_tree = template_slide._element.spTree
        for child in orig_sp_tree:
            src_sp_tree.append(copy.deepcopy(child))

        # 이미지 관계(rId) 복사: 템플릿의 미디어를 새 슬라이드에 등록
        for rel in template_slide.part.rels.values():
            if "image" in rel.reltype:
                img_part = rel.target_part
                new_slide.part.relate_to(img_part, rel.reltype)

        # 텍스트 치환
        fill_slide(new_slide, mapping)

    # 바이트로 저장
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
                ppt_bytes = ppt_file.read()
                prs_template = Presentation(io.BytesIO(ppt_bytes))

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
