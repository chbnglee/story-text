"""
Story Tool — Streamlit Web App
  Step 1. 음원 생성   : Excel/CSV 업로드 → Normal / Easy / Difficult / mBook xlsx 생성
  Step 2. ICMS txt 생성: xlsx 3종 업로드 → 감정 태그 포함 TXT 생성 (Gemini API 사용)
"""

import io
import json
import re
import zipfile
from collections import defaultdict

import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ─────────────────────────────────────────────────────────────────────────────
# 공통 상수 & 스타일
# ─────────────────────────────────────────────────────────────────────────────

LEVELS = ["Normal", "Easy", "Difficult"]

HEADER_FILL = PatternFill("solid", fgColor="4F81BD")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
CELL_FONT   = Font(name="Calibri", size=11)
WRAP_ALIGN  = Alignment(wrap_text=True, vertical="top")
THIN_SIDE   = Side(style="thin", color="BFBFBF")
THIN_BORDER = Border(
    left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE
)
COL_WIDTHS = {"ID": 12, "Key": 22, "Text": 80}

SCENE_RE = re.compile(r"#SC(\d+)", re.MULTILINE)

EMOTION_PROMPT = """\
You are tagging the emotional tone of each scene in a children's story.

Available tags: Neutral, Happy, Sad, Angry
Rules:
- Every scene value MUST start with "Neutral".
- You MUST assign at least one non-Neutral tag (Happy, Sad, or Angry) to AT LEAST 3 scenes.
- Look for: crying / loss / disappointment → Sad | cheering / success / reunion → Happy | frustration / conflict → Angry
- Multiple tags are allowed (e.g. "Neutral, Sad, Happy" for bittersweet moments).
- Err toward tagging — a subtle emotional cue is enough.

Return ONLY a JSON object. Keys = scene IDs, values = comma-separated tag strings.
Example:
{
  "SC01": "Neutral",
  "SC02": "Neutral, Sad",
  "SC03": "Neutral",
  "SC04": "Neutral, Happy",
  "SC05": "Neutral, Angry, Sad"
}

Story scenes (Normal version):
"""

# ─────────────────────────────────────────────────────────────────────────────
# 공통 유틸
# ─────────────────────────────────────────────────────────────────────────────

def safe_fname(s: str) -> str:
    """파일명에 쓸 수 없는 문자 제거."""
    return re.sub(r'[\\/*?:"<>|]', "_", s)


def _style_ws(ws):
    """헤더·데이터 행 스타일 일괄 적용."""
    for cell in ws[1]:
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.border    = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font      = CELL_FONT
            cell.border    = THIN_BORDER
            cell.alignment = WRAP_ALIGN
    for idx, col in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[
            openpyxl.utils.get_column_letter(idx)
        ].width = COL_WIDTHS[col]
    ws.row_dimensions[1].height = 18


# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 로직
# ─────────────────────────────────────────────────────────────────────────────

def split_sentences(text: str) -> list[str]:
    raw = re.split(r'(?<=[.!?"])\s+', text.strip())
    return [s.strip() for s in raw if s.strip()]


def parse_scenes(raw: str) -> list[tuple[int, list[str]]]:
    parts = SCENE_RE.split(raw.strip())
    result, i = [], 1
    while i < len(parts) - 1:
        sc_num     = int(parts[i])
        sc_text    = parts[i + 1].strip()
        sentences  = split_sentences(sc_text)
        if sentences:
            result.append((sc_num, sentences))
        i += 2
    return result


def make_level_xlsx(story_id: str, level_text: str, initial: str) -> bytes:
    """Normal / Easy / Difficult 중 하나의 xlsx 바이트 반환."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = initial
    ws.append(["ID", "Key", "Text"])
    for sc_num, sentences in parse_scenes(level_text):
        for si, sent in enumerate(sentences, start=1):
            ws.append([story_id, f"SC{sc_num:02d}_ST{si:02d}_{initial}", sent])
    _style_ws(ws)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def make_mbook_xlsx(story_id: str, mbook_text: str) -> bytes:
    """mBook 전체 텍스트를 1개 셀에 넣은 xlsx 바이트 반환."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "mBook"
    ws.append(["ID", "Key", "Text"])
    ws.append([story_id, "mBook_Script", mbook_text])
    _style_ws(ws)
    # 텍스트 행 높이: 줄 수에 비례
    line_count = max(mbook_text.count("\n") + 1, 1)
    ws.row_dimensions[2].height = max(line_count * 15, 30)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def read_input_file(file_bytes: bytes, filename: str) -> list[dict]:
    """xlsx 또는 csv 업로드를 읽어 dict 목록 반환."""
    if filename.lower().endswith(".csv"):
        import csv, io as _io
        for enc in ("utf-8-sig", "utf-8", "cp949", "latin-1"):
            try:
                text = file_bytes.decode(enc)
                reader = csv.DictReader(_io.StringIO(text))
                return [dict(r) for r in reader]
            except (UnicodeDecodeError, LookupError):
                continue
        raise ValueError("CSV 인코딩을 자동으로 감지할 수 없습니다.")
    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        ws = wb.active
        headers = [
            str(c.value).strip() if c.value is not None else f"col{i}"
            for i, c in enumerate(ws[1])
        ]
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if all(v is None for v in row):
                continue
            rows.append({
                headers[i]: (str(row[i]) if row[i] is not None else "")
                for i in range(len(headers))
            })
        return rows


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 로직
# ─────────────────────────────────────────────────────────────────────────────

def read_xlsx_rows(file_bytes: bytes) -> list[tuple[str, str, str]]:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        rows.append((str(row[0]), str(row[1]), str(row[2])))
    return rows


def extract_scene_texts(rows: list[tuple[str, str, str]]) -> dict[str, str]:
    scenes: dict[str, list[str]] = defaultdict(list)
    for _, key, text in rows:
        scenes[key.split("_")[0]].append(text)
    return {sc: " ".join(sents) for sc, sents in scenes.items()}


def _gemini_call(client, model: str, prompt: str) -> str:
    return client.models.generate_content(model=model, contents=prompt).text.strip()


def get_emotions(client, model: str, scene_texts: dict[str, str]) -> dict[str, str]:
    block = "\n\n".join(f"{k}:\n{v}" for k, v in sorted(scene_texts.items()))
    raw   = _gemini_call(client, model, EMOTION_PROMPT + block)

    def parse(r: str) -> dict:
        m = re.search(r"\{[\s\S]*\}", r)
        if not m:
            raise ValueError("JSON을 찾을 수 없습니다.")
        return json.loads(m.group())

    data = parse(raw)
    non_neutral = [k for k, v in data.items() if v.strip() != "Neutral"]
    if len(non_neutral) < 3:
        retry = (
            f"Your previous answer tagged only {len(non_neutral)} scene(s) as non-Neutral. "
            f"The requirement is AT LEAST 3 scenes with Happy, Sad, or Angry.\n"
            f"Go through every scene again and find emotional cues. "
            f"Return the complete JSON with at least 3 non-Neutral tags.\n\n"
            f"Story scenes:\n{block}"
        )
        data = parse(_gemini_call(client, model, retry))
    return data


def build_txt_content(
    level_rows: dict[str, list[tuple[str, str, str]]],
    emotions: dict[str, str],
    scene_keys: list[str],
) -> str:
    lines = []
    for level in LEVELS:
        for _, key, text in level_rows.get(level, []):
            lines.append(f"{key} = {text}")
        lines.append("")
    for sc in sorted(scene_keys):
        lines.append(f"{sc}_Emotion = {emotions.get(sc, 'Neutral')}")
    return "\n".join(lines)


def group_xlsx_files(
    files: list,
) -> dict[str, dict[str, bytes]]:
    """업로드된 파일들을 기본 이름(base)과 레벨로 그룹화."""
    pat = re.compile(r"^(.+?)_(Normal|Easy|Difficult)\.xlsx$", re.IGNORECASE)
    groups: dict[str, dict[str, bytes]] = defaultdict(dict)
    unmatched = []
    for f in files:
        m = pat.match(f.name)
        if m:
            base  = m.group(1)
            level = m.group(2).capitalize()
            groups[base][level] = f.read()
        else:
            unmatched.append(f.name)
    return dict(groups), unmatched


# ─────────────────────────────────────────────────────────────────────────────
# 페이지 설정
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Story Tool",
    page_icon="📖",
    layout="centered",
)

st.markdown("""
<style>
/* 탭 폰트 강조 */
.stTabs [data-baseweb="tab"] { font-size: 1.05rem; font-weight: 600; padding: 8px 24px; }
</style>
""", unsafe_allow_html=True)

# 세션 상태 초기화
for key in ("step1_zip", "step2_result"):
    if key not in st.session_state:
        st.session_state[key] = None

tab1, tab2 = st.tabs(["Step 1. 음원 생성", "Step 2. ICMS용 txt 생성"])


# ═════════════════════════════════════════════════════════
# STEP 1 — 음원 생성
# ═════════════════════════════════════════════════════════
with tab1:
    st.subheader("Step 1. 음원 생성")
    st.caption(
        "**ID / Title / Normal / Easy / Difficult / mBook** 열이 있는 "
        "Excel(.xlsx) 또는 CSV 파일을 업로드하세요."
    )

    st.markdown("""
<div style="background:#e8f4fd; border-left:4px solid #4F81BD;
            padding:10px 14px; border-radius:4px; font-size:0.82rem; line-height:1.7;">
📌 Normal / Easy / Difficult 열: #SC01 기준으로 문장 단위로 분할됩니다.<br>
📌 mBook 열: 전체 텍스트가 한 셀에 그대로 저장됩니다.
</div>
""", unsafe_allow_html=True)

    uploaded_input = st.file_uploader(
        "파일 선택 (.xlsx / .csv)",
        type=["xlsx", "csv"],
        key="step1_file",
    )

    col_run, col_dl = st.columns([1, 1])

    with col_run:
        run_btn = st.button("📂 xlsx 생성", use_container_width=True, key="step1_run")

    if run_btn and uploaded_input:
        with st.spinner("처리 중..."):
            try:
                raw_bytes = uploaded_input.read()
                rows = read_input_file(raw_bytes, uploaded_input.name)

                zip_buf = io.BytesIO()
                counts  = {"Normal": 0, "Easy": 0, "Difficult": 0, "mBook": 0}

                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for row in rows:
                        sid   = row.get("ID", "").strip()
                        title = row.get("Title", "").strip()
                        if not sid or not title:
                            continue
                        st_title = safe_fname(title)

                        for col_name, initial in [
                            ("Normal", "N"), ("Easy", "E"), ("Difficult", "D")
                        ]:
                            raw = row.get(col_name, "").strip()
                            if raw:
                                data = make_level_xlsx(sid, raw, initial)
                                zf.writestr(f"{sid}_{st_title}_{col_name}.xlsx", data)
                                counts[col_name] += 1

                        mbook = row.get("mBook", "").strip()
                        if mbook:
                            data = make_mbook_xlsx(sid, mbook)
                            zf.writestr(f"{sid}_{st_title}_mBook.xlsx", data)
                            counts["mBook"] += 1

                zip_buf.seek(0)
                st.session_state.step1_zip = zip_buf.getvalue()

                total = sum(counts.values())
                st.success(
                    f"✅ 완료! {len(rows)}개 스토리 → "
                    f"Normal {counts['Normal']}개 / Easy {counts['Easy']}개 / "
                    f"Difficult {counts['Difficult']}개 / mBook {counts['mBook']}개 "
                    f"(총 {total}개 파일)"
                )

            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")
                st.session_state.step1_zip = None

    if st.session_state.step1_zip:
        with col_dl:
            st.download_button(
                label="⬇️ ZIP 다운로드",
                data=st.session_state.step1_zip,
                file_name="script_output.zip",
                mime="application/zip",
                use_container_width=True,
                key="step1_dl",
            )


# ═════════════════════════════════════════════════════════
# STEP 2 — ICMS용 txt 생성
# ═════════════════════════════════════════════════════════
with tab2:
    st.subheader("Step 2. ICMS용 txt 생성")
    st.caption(
        "**ID_Title_Normal.xlsx / ID_Title_Easy.xlsx / ID_Title_Difficult.xlsx** "
        "파일을 업로드하면 감정 태그가 포함된 TXT 파일을 생성합니다. "
        "여러 스토리를 동시에 처리할 수 있습니다."
    )

    # ── API 키 ──────────────────────────────────────────
    _secret_key = ""
    try:
        _secret_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        pass

    if _secret_key:
        api_key = _secret_key
        st.success("✅ Gemini API 키가 설정되어 있습니다.", icon="🔑")
    else:
        api_key = st.text_input(
            "🔑 Gemini API Key",
            type="password",
            placeholder="AIza...",
            help="https://aistudio.google.com/app/apikey 에서 무료 발급",
        )

    # ── 파일 업로드 ──────────────────────────────────────
    uploaded_xlsx = st.file_uploader(
        "xlsx 파일 선택 (Normal + Easy + Difficult, 여러 스토리 동시 가능)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="step2_files",
    )

    col_run2, col_dl2 = st.columns([1, 1])

    with col_run2:
        run_btn2 = st.button("📄 TXT 생성", use_container_width=True, key="step2_run")

    if run_btn2:
        if not api_key:
            st.warning("⚠️ Gemini API 키를 먼저 입력해주세요.")
        elif not uploaded_xlsx:
            st.warning("⚠️ xlsx 파일을 업로드해주세요.")
        else:
            with st.spinner("Gemini로 감정 분석 중... 잠시 기다려주세요."):
                try:
                    from google import genai as _genai

                    client = _genai.Client(api_key=api_key)
                    MODEL  = "gemini-2.5-flash"

                    groups, unmatched = group_xlsx_files(uploaded_xlsx)

                    if unmatched:
                        st.warning(
                            f"파일명 형식을 인식할 수 없어 건너뜀: {', '.join(unmatched)}\n"
                            "파일명은 **ID_Title_Normal.xlsx** 형식이어야 합니다."
                        )
                    if not groups:
                        st.error("인식된 파일 그룹이 없습니다. 파일명을 확인해주세요.")
                        st.stop()

                    results: dict[str, bytes] = {}
                    prog = st.progress(0, text="처리 중...")

                    for i, (base, level_files) in enumerate(sorted(groups.items())):
                        prog.progress(
                            (i) / len(groups),
                            text=f"[{i+1}/{len(groups)}] {base} 처리 중...",
                        )

                        level_rows = {
                            lvl: read_xlsx_rows(data)
                            for lvl, data in level_files.items()
                            if lvl in LEVELS
                        }

                        normal_rows = level_rows.get("Normal", [])
                        scene_texts = extract_scene_texts(normal_rows)
                        scene_keys  = sorted(scene_texts.keys())

                        emotions = get_emotions(client, MODEL, scene_texts)
                        txt_str  = build_txt_content(level_rows, emotions, scene_keys)

                        parts    = base.split("_", 1)
                        sid      = parts[0]
                        title    = parts[1] if len(parts) > 1 else base
                        out_name = f"{sid}_{title}.txt"
                        results[out_name] = txt_str.encode("utf-8")

                    prog.progress(1.0, text="완료!")

                    st.session_state.step2_result = results
                    st.success(f"✅ {len(results)}개 TXT 파일 생성 완료!")

                except Exception as e:
                    st.error(f"오류가 발생했습니다: {e}")
                    st.exception(e)
                    st.session_state.step2_result = None

    if st.session_state.step2_result:
        results = st.session_state.step2_result
        with col_dl2:
            if len(results) == 1:
                fname, content = next(iter(results.items()))
                st.download_button(
                    label=f"⬇️ {fname}",
                    data=content,
                    file_name=fname,
                    mime="text/plain",
                    use_container_width=True,
                    key="step2_dl",
                )
            else:
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fname, content in results.items():
                        zf.writestr(fname, content)
                zip_buf.seek(0)
                st.download_button(
                    label="⬇️ ZIP 다운로드",
                    data=zip_buf.getvalue(),
                    file_name="txt_output.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key="step2_dl",
                )
