import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import openpyxl
import re

# ── NEIS 형식 자동 파싱 ──────────────────────────────────────
def parse_neis_excel(file_bytes):
    """
    NEIS 엑셀에서 반/번호·성명·점수·결시명칭만 추출.
    - 반/번호: B열(index 1)의 'N/N' 패턴
    - 성명:    B열 이후 첫 번째 문자열 열
    - 점수:    0~9999 범위 숫자 열 (학번 같은 큰 수 자동 제외)
    - 결시명칭: 성명·점수 열 이외의 문자열 열
    반환: (students_list, score_col_count, error_msg)
    """
    wb = openpyxl.load_workbook(BytesIO(file_bytes))
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))

    STOP_KEYWORDS = {"합계", "참가", "인원", "평균"}

    # 1. 파일 종류 판별
    #    정기시험: row8(index7) B열(col1)에 '만점' 문자열 존재
    #    수행평가: row8(index7) B열이 비어 있음
    row8_b = all_rows[7][1] if len(all_rows) > 7 and len(all_rows[7]) > 1 else None
    is_perf = not bool(row8_b and str(row8_b).strip())

    # 정기시험이면 row7(index6) B열에서 '중간'/'기말' 키워드로 시험 종류 판별
    exam_type = None
    if not is_perf:
        row7_b = all_rows[6][1] if len(all_rows) > 6 and len(all_rows[6]) > 1 else ""
        row7_b = row7_b or ""
        if "중간" in row7_b:
            exam_type = "중간고사"
        elif "기말" in row7_b:
            exam_type = "기말고사"

    # 2. 학생 행: B열이 'N/N' 패턴
    student_indices = []
    for i, row in enumerate(all_rows):
        val = row[1] if len(row) > 1 else None
        if val and re.match(r'^\d+/\d+$', str(val).strip()):
            student_indices.append(i)

    if not student_indices:
        return None, None, None, "학생 행(반/번호)을 찾을 수 없습니다."

    # 합계·인원 행 이후 제거
    cleaned = []
    for i in student_indices:
        if any(v and any(k in str(v) for k in STOP_KEYWORDS) for v in all_rows[i]):
            break
        cleaned.append(i)
    student_indices = cleaned

    # 3. 열 역할 결정 (전체 학생 행 기준)
    name_col   = None
    score_cols = []
    note_cols  = []

    for j in range(2, 20):
        vals = [all_rows[i][j] for i in student_indices if j < len(all_rows[i])]
        non_empty = [v for v in vals if v is not None and str(v).strip()]
        if not non_empty:
            continue
        numeric = [v for v in non_empty if isinstance(v, (int, float)) and not isinstance(v, bool)]
        strings = [v for v in non_empty if isinstance(v, str) and str(v).strip()]

        if name_col is None and strings:
            name_col = j
        elif numeric and all(0 <= v <= 9999 for v in numeric):
            score_cols.append(j)
        elif strings and name_col is not None:
            note_cols.append(j)

    if name_col is None:
        return None, None, None, "성명 열을 찾을 수 없습니다."
    if not score_cols:
        return None, None, None, "점수 열을 찾을 수 없습니다."

    # 4. 수행평가 합계 열 자동 제외
    #    합계 열: 다른 점수 열들의 합과 거의 일치하는 마지막 열
    if is_perf and len(score_cols) >= 3:
        part_cols = score_cols[:-1]
        sum_col   = score_cols[-1]
        match_count = 0
        total_valid = 0
        for i in student_indices:
            row = all_rows[i]
            parts = [row[c] for c in part_cols if c < len(row) and isinstance(row[c], (int, float))]
            total = row[sum_col] if sum_col < len(row) else None
            if parts and isinstance(total, (int, float)):
                total_valid += 1
                if abs(sum(parts) - total) < 0.1:
                    match_count += 1
        if total_valid > 0 and match_count / total_valid >= 0.8:
            score_cols = part_cols  # 합계 열 제외

    # 5. 데이터 추출
    students = []
    for i in student_indices:
        row = all_rows[i]
        no   = str(row[1]).strip()
        name = str(row[name_col]).strip() if row[name_col] else ""
        scores = [
            row[sc] if sc < len(row) and isinstance(row[sc], (int, float)) and not isinstance(row[sc], bool) else None
            for sc in score_cols
        ]
        note = next(
            (str(row[nc]).strip() for nc in note_cols if nc < len(row) and isinstance(row[nc], str) and str(row[nc]).strip()),
            ""
        )
        students.append({"번호": no, "이름": name, "scores": scores, "note": note})

    return students, len(score_cols), is_perf, exam_type, None

st.set_page_config(page_title="학기말 성적 산출", page_icon="📊", layout="wide")

# ── 성취도 등급 ──────────────────────────────────────────────
def achievement_grade(score):
    if pd.isna(score):
        return "-"
    if score >= 90: return "A"
    if score >= 80: return "B"
    if score >= 70: return "C"
    if score >= 60: return "D"
    return "E"

# ── 석차 등급 (한국 9등급제) ─────────────────────────────────
def rank_grade(rank, total):
    pct = rank / total * 100
    if pct <= 10:  return 1
    if pct <= 34:  return 2
    if pct <= 66:  return 3
    if pct <= 90:  return 4
    return 5

# ── 세션 상태 초기화 ─────────────────────────────────────────
def init_state():
    defaults = {
        "subject": "국어",
        "grade": "1",
        "semester": "1",
        "class_no": "1",
        "mid_ratio": 35.0,
        "final_ratio": 35.0,
        "perf_count": 2,
        "perf_names": ["수행평가1", "수행평가2"],
        "perf_ratios": [15.0, 15.0],
        "perf_max": [100.0, 100.0],
        "mid_max": 100.0,
        "final_max": 100.0,
        "students": pd.DataFrame(columns=["번호", "이름"]),
        "calculated": False,
        "result_df": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ── 제목 ─────────────────────────────────────────────────────
st.title("📊 학기말 성적 산출 시스템")
st.caption("정기시험(중간·기말) + 수행평가 반영 비율에 따른 최종 성적 산출")

tab1, tab2, tab3, tab4 = st.tabs(["⚙️ 기본 설정", "✏️ 점수 입력", "📋 성적 산출", "📈 대시보드"])

# ═══════════════════════════════════════════════════════════════
# TAB 1 : 기본 설정
# ═══════════════════════════════════════════════════════════════
with tab1:
    st.subheader("수업 정보")
    c1, c2, c3, c4 = st.columns(4)
    st.session_state.subject  = c1.text_input("과목명", st.session_state.subject)
    st.session_state.grade    = c2.selectbox("학년", ["1","2","3"], index=int(st.session_state.grade)-1)
    st.session_state.semester = c3.selectbox("학기", ["1","2"], index=int(st.session_state.semester)-1)
    st.session_state.class_no = c4.text_input("반", st.session_state.class_no)

    st.divider()
    st.subheader("반영 비율 설정")
    st.caption("모든 비율의 합이 **100%** 가 되도록 설정하세요.")

    c1, c2 = st.columns(2)
    st.session_state.mid_ratio   = c1.number_input("중간고사 비율 (%)", 0.0, 100.0, st.session_state.mid_ratio, 1.0)
    st.session_state.final_ratio = c2.number_input("기말고사 비율 (%)", 0.0, 100.0, st.session_state.final_ratio, 1.0)

    st.session_state.mid_max   = c1.number_input("중간고사 만점", 1.0, 1000.0, st.session_state.mid_max, 1.0)
    st.session_state.final_max = c2.number_input("기말고사 만점", 1.0, 1000.0, st.session_state.final_max, 1.0)

    st.markdown("**수행평가 항목**")
    st.session_state.perf_count = st.number_input(
        "수행평가 항목 수", 1, 10, st.session_state.perf_count, 1
    )
    n = st.session_state.perf_count

    # 항목 목록 길이 조정
    for lst, default in [
        ("perf_names", [f"수행평가{i+1}" for i in range(n)]),
        ("perf_ratios", [10.0]*n),
        ("perf_max",    [100.0]*n),
    ]:
        cur = st.session_state[lst]
        if len(cur) < n:
            st.session_state[lst] = cur + default[len(cur):]
        elif len(cur) > n:
            st.session_state[lst] = cur[:n]

    cols = st.columns(3)
    cols[0].markdown("항목 이름")
    cols[1].markdown("반영 비율 (%)")
    cols[2].markdown("만점")

    for i in range(n):
        c1, c2, c3 = st.columns(3)
        st.session_state.perf_names[i]  = c1.text_input(f"이름_{i}", st.session_state.perf_names[i], label_visibility="collapsed", key=f"pname_{i}")
        st.session_state.perf_ratios[i] = c2.number_input(f"비율_{i}", 0.0, 100.0, float(st.session_state.perf_ratios[i]), 1.0, label_visibility="collapsed", key=f"pratio_{i}")
        st.session_state.perf_max[i]    = c3.number_input(f"만점_{i}", 1.0, 1000.0, float(st.session_state.perf_max[i]), 1.0, label_visibility="collapsed", key=f"pmax_{i}")

    total_ratio = st.session_state.mid_ratio + st.session_state.final_ratio + sum(st.session_state.perf_ratios[:n])
    if abs(total_ratio - 100) < 0.01:
        st.success(f"✅ 반영 비율 합계: {total_ratio:.1f}%")
    else:
        st.error(f"⚠️ 반영 비율 합계: {total_ratio:.1f}% (100%가 되어야 합니다)")


# ═══════════════════════════════════════════════════════════════
# TAB 2 : 점수 입력
# ═══════════════════════════════════════════════════════════════
with tab2:
    st.subheader("학생 점수 입력")

    # ── Excel 다중 업로드 ────────────────────────────────────
    with st.expander("📂 Excel 파일 여러 개 업로드 (NEIS 형식 자동 인식)", expanded=True):
        st.caption("중간고사·기말고사·수행평가 파일을 한 번에 선택하세요. 번/학번 형식(1/1, 1/2…)이 있는 NEIS 파일은 자동으로 파싱됩니다.")

        uploaded_files = st.file_uploader(
            "Excel 파일 업로드 (.xlsx) — 여러 개 동시 선택 가능",
            type=["xlsx"],
            accept_multiple_files=True,
        )

        if uploaded_files:
            n_perf_cur = st.session_state.perf_count
            perf_options = ["중간고사", "기말고사"] + [st.session_state.perf_names[j] for j in range(n_perf_cur)]

            # ── 1단계: 모든 파일 파싱 후 그룹 분류 ─────────────
            groups = {"중간고사": [], "기말고사": [], "수행평가": [], "알수없음": []}
            file_cache = {}  # uf.name → (students, n_scores, is_perf, exam_type)

            for uf in uploaded_files:
                file_bytes = uf.read()
                students, n_scores, is_perf, exam_type, err = parse_neis_excel(file_bytes)
                if err or not students:
                    groups["알수없음"].append(uf.name)
                    continue
                file_cache[uf.name] = (students, n_scores, is_perf, exam_type)
                if is_perf:
                    groups["수행평가"].append(uf.name)
                elif exam_type == "중간고사":
                    groups["중간고사"].append(uf.name)
                elif exam_type == "기말고사":
                    groups["기말고사"].append(uf.name)
                else:
                    groups["알수없음"].append(uf.name)

            # ── 2단계: 그룹별로 묶어서 표시 ─────────────────────
            parsed_frames = {}

            group_labels = {
                "중간고사": "📄 정기시험 — 중간고사",
                "기말고사": "📄 정기시험 — 기말고사",
                "수행평가": "📝 수행평가",
                "알수없음": "❓ 알수없음",
            }

            for group_key, fnames in groups.items():
                if not fnames:
                    continue
                st.markdown(f"---\n### {group_labels[group_key]} ({len(fnames)}개 파일)")

                for fname in fnames:
                    if fname not in file_cache:
                        st.warning(f"{fname}: 파싱 실패")
                        continue
                    students, n_scores, is_perf, exam_type = file_cache[fname]

                    with st.expander(f"📂 {fname}  ·  {len(students)}명", expanded=False):
                        # 드롭다운
                        type_choices = []
                        cols_ui = st.columns(n_scores)
                        for k in range(n_scores):
                            if is_perf:
                                default_idx = min(2 + k, len(perf_options) - 1)
                            elif exam_type == "중간고사":
                                default_idx = 0
                            elif exam_type == "기말고사":
                                default_idx = 1
                            else:
                                default_idx = min(k, 1)
                            choice = cols_ui[k].selectbox(
                                f"점수 열 {k+1} 종류",
                                perf_options,
                                index=default_idx,
                                key=f"stype_{fname}_{k}",
                            )
                            type_choices.append(choice)

                        # 미리보기
                        preview_rows = []
                        for s in students:
                            row = {"번호": s["번호"], "이름": s["이름"]}
                            for k, tc in enumerate(type_choices):
                                row[tc] = s["scores"][k] if k < len(s["scores"]) else None
                            row["비고"] = s["note"]
                            preview_rows.append(row)
                        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True, height=280)

                    # 파싱 결과 누적 (expander 밖에서)
                    for s in students:
                        name = s["이름"]
                        # 반/번 저장 (번호 "3/5" → 반=3, 번=5)
                        parsed_frames.setdefault("__bango__", {})[name] = s["번호"]
                        for k, tc in enumerate(type_choices):
                            parsed_frames.setdefault(tc, {})[name] = s["scores"][k] if k < len(s["scores"]) else None
                        if s["note"]:
                            # 어느 고사에 결시인지 기록 (tc = 점수 종류)
                            note_key = f"__note_{type_choices[0]}__"
                            parsed_frames.setdefault(note_key, {})[name] = s["note"]
                            parsed_frames.setdefault("__note__", {})[name] = s["note"]

            if parsed_frames and st.button("⬆️ 업로드 파일로 점수 불러오기", type="primary"):
                # ── 학생 목록 수집 및 반별 정렬 ──────────────────
                all_names = []
                for k, v in parsed_frames.items():
                    if not k.startswith("__"):
                        all_names += list(v.keys())
                all_names = list(dict.fromkeys(all_names))

                def sort_key(name):
                    bango = parsed_frames.get("__bango__", {}).get(name, "0/0")
                    parts = str(bango).split("/")
                    try:
                        return (int(parts[0]), int(parts[1]))
                    except Exception:
                        return (0, 0)

                all_names.sort(key=sort_key)

                # ── 기본 점수 수집 ────────────────────────────────
                rows = []
                for name in all_names:
                    bango = parsed_frames.get("__bango__", {}).get(name, "")
                    parts = str(bango).split("/")
                    ban = parts[0] if len(parts) == 2 else ""
                    bun = parts[1] if len(parts) == 2 else ""
                    row = {
                        "반": ban,
                        "번": bun,
                        "이름": name,
                        "중간고사": parsed_frames.get("중간고사", {}).get(name),
                        "기말고사": parsed_frames.get("기말고사", {}).get(name),
                    }
                    for j in range(st.session_state.perf_count):
                        pname = st.session_state.perf_names[j]
                        row[pname] = parsed_frames.get(pname, {}).get(name)
                    row["비고"] = parsed_frames.get("__note__", {}).get(name, "")
                    rows.append(row)

                df_merge = pd.DataFrame(rows)

                # ── 결시 점수 처리 ────────────────────────────────
                exam_cols = ["중간고사", "기말고사"]

                for exam in exam_cols:
                    note_key = f"__note_{exam}__"
                    note_map = parsed_frames.get(note_key, {})
                    if not note_map:
                        continue

                    # 해당 고사의 유효 점수 목록 (결시 제외)
                    valid_scores = df_merge.loc[
                        ~df_merge["이름"].isin(note_map.keys()),
                        exam
                    ].dropna().tolist()
                    min_score = min(valid_scores) if valid_scores else 0

                    for name, note in note_map.items():
                        idx = df_merge[df_merge["이름"] == name].index
                        if idx.empty:
                            continue

                        if "미인정" in note:
                            # 차하점 = 최저점 - 1 (최소 0)
                            chahajum = max(0, min_score - 1)
                            df_merge.loc[idx, exam] = chahajum
                            df_merge.loc[idx, "비고"] = f"{note}→차하점({chahajum}점)"

                        elif "질병" in note:
                            # 다른 고사 점수의 80%
                            other_exam = "기말고사" if exam == "중간고사" else "중간고사"
                            other_score = df_merge.loc[idx, other_exam].values[0]
                            if other_score is not None and not pd.isna(other_score):
                                calc = round(float(other_score) * 0.8, 1)
                                df_merge.loc[idx, exam] = calc
                                df_merge.loc[idx, "비고"] = f"{note}→{other_exam}×80%({calc}점)"

                st.session_state["uploaded_df"] = df_merge
                st.success(f"✅ {len(all_names)}명 데이터 불러오기 완료! (반별 정렬 적용)")
                st.rerun()

    # ── 학생 수 설정 ─────────────────────────────────────────
    if "uploaded_df" in st.session_state and not st.session_state["uploaded_df"].empty:
        n_students = len(st.session_state["uploaded_df"])
    else:
        n_students = st.number_input("학생 수", 1, 100, 16, 1)
    n_perf = st.session_state.perf_count

    # ── 점수 입력 테이블 구성 ─────────────────────────────────
    if "uploaded_df" in st.session_state:
        base = st.session_state["uploaded_df"]
    else:
        base = pd.DataFrame()

    def get_col(df, col, idx, default=""):
        if df.empty or col not in df.columns or idx >= len(df):
            return default
        v = df[col].iloc[idx]
        return "" if pd.isna(v) else v

    rows = []
    has_ban = "반" in base.columns if not base.empty else False
    for i in range(n_students):
        row = {
            "반": get_col(base, "반", i, ""),
            "번": get_col(base, "번", i, ""),
            "이름": get_col(base, "이름", i, ""),
            "중간고사": get_col(base, "중간고사", i, ""),
            "기말고사": get_col(base, "기말고사", i, ""),
        }
        for j in range(n_perf):
            col_name = st.session_state.perf_names[j]
            row[col_name] = get_col(base, col_name, i, "")
        row["비고"] = get_col(base, "비고", i, "")
        rows.append(row)

    df_input = pd.DataFrame(rows)

    edited = st.data_editor(
        df_input,
        use_container_width=True,
        num_rows="fixed",
        key="score_editor",
        column_config={
            "반": st.column_config.TextColumn("반"),
            "번": st.column_config.TextColumn("번"),
            "이름": st.column_config.TextColumn("이름"),
            "중간고사": st.column_config.NumberColumn(f"중간고사 (/{st.session_state.mid_max:.0f}점)"),
            "기말고사": st.column_config.NumberColumn(f"기말고사 (/{st.session_state.final_max:.0f}점)"),
            **{
                st.session_state.perf_names[j]: st.column_config.NumberColumn(
                    f"{st.session_state.perf_names[j]} (/{st.session_state.perf_max[j]:.0f}점)"
                )
                for j in range(n_perf)
            },
            "비고": st.column_config.TextColumn("비고"),
        },
    )

    if st.button("🧮 성적 산출하기", type="primary", use_container_width=True):
        # ── 환산 점수 계산 ──────────────────────────────────
        df = edited.copy()
        df["중간고사"] = pd.to_numeric(df["중간고사"], errors="coerce")
        df["기말고사"] = pd.to_numeric(df["기말고사"], errors="coerce")

        mid_r   = st.session_state.mid_ratio / 100
        final_r = st.session_state.final_ratio / 100
        mid_max   = st.session_state.mid_max
        final_max = st.session_state.final_max

        df["중간환산"] = (df["중간고사"] / mid_max * mid_r * 100).round(2)
        df["기말환산"] = (df["기말고사"] / final_max * final_r * 100).round(2)

        for j in range(n_perf):
            col = st.session_state.perf_names[j]
            df[col] = pd.to_numeric(df[col], errors="coerce")
            r = st.session_state.perf_ratios[j] / 100
            mx = st.session_state.perf_max[j]
            df[f"{col}_환산"] = (df[col] / mx * r * 100).round(2)

        perf_conv_cols = [f"{st.session_state.perf_names[j]}_환산" for j in range(n_perf)]
        df["최종점수"] = (df["중간환산"].fillna(0) + df["기말환산"].fillna(0)
                        + df[perf_conv_cols].fillna(0).sum(axis=1)).round(2)

        # 비고 처리
        if "비고" in df.columns:
            bigo = df["비고"].astype(str)
            # 자퇴생: 석차/성취도 산출 대상에서 완전 제외
            mask_quit = bigo.str.contains("자퇴", na=False)
            df.loc[mask_quit, "최종점수"] = np.nan
            # 결시(미인정결시·질병결 등): 최종점수 NaN 처리
            # 업로드 시 이미 처리된 경우(비고에 '→' 포함)는 제외, 미처리 결시만 NaN
            mask_absent = (
                bigo.str.contains("결시|결석|질병결", na=False) &
                ~mask_quit &
                ~bigo.str.contains("→", na=False)
            )
            df.loc[mask_absent, "최종점수"] = np.nan

        # 석차: 자퇴·결시 제외한 학생만 산출
        valid = df["최종점수"].notna()
        df["석차"] = np.nan
        if valid.sum() > 0:
            df.loc[valid, "석차"] = df.loc[valid, "최종점수"].rank(ascending=False, method="min").astype(int)

        df["성취도"] = df["최종점수"].apply(achievement_grade)

        total_valid = valid.sum()
        if total_valid > 0:
            df["석차등급"] = df["석차"].apply(
                lambda r: rank_grade(r, total_valid) if pd.notna(r) else "-"
            )
        else:
            df["석차등급"] = "-"

        # 자퇴생 명시적 표시
        if "비고" in df.columns:
            quit_mask = df["비고"].astype(str).str.contains("자퇴", na=False)
            df.loc[quit_mask, "성취도"]  = "자퇴"
            df.loc[quit_mask, "석차등급"] = "자퇴"

        st.session_state.result_df = df
        st.session_state.calculated = True
        st.success("성적 산출 완료! '성적 산출' 탭에서 결과를 확인하세요.")


# ═══════════════════════════════════════════════════════════════
# TAB 3 : 성적 산출
# ═══════════════════════════════════════════════════════════════
with tab3:
    if not st.session_state.calculated or st.session_state.result_df is None:
        st.info("'점수 입력' 탭에서 점수를 입력하고 **성적 산출하기** 버튼을 눌러주세요.")
    else:
        df = st.session_state.result_df
        n_perf = st.session_state.perf_count

        # 헤더 정보
        st.markdown(f"""
**{st.session_state.grade}학년 {st.session_state.semester}학기 {st.session_state.class_no}반 | 과목: {st.session_state.subject}**
반영 비율 — 중간고사: {st.session_state.mid_ratio}% / 기말고사: {st.session_state.final_ratio}% / 수행평가: {sum(st.session_state.perf_ratios[:n_perf])}%
        """)

        # 요약 지표
        valid_scores = df["최종점수"].dropna()
        mc1, mc2, mc3, mc4, mc5 = st.columns(5)
        mc1.metric("응시 인원", f"{len(valid_scores)}명")
        mc2.metric("평균", f"{valid_scores.mean():.2f}점" if len(valid_scores) else "-")
        mc3.metric("최고", f"{valid_scores.max():.2f}점" if len(valid_scores) else "-")
        mc4.metric("최저", f"{valid_scores.min():.2f}점" if len(valid_scores) else "-")
        mc5.metric("표준편차", f"{valid_scores.std():.2f}" if len(valid_scores) else "-")

        st.divider()

        # 결과 테이블 표시 컬럼 선택
        show_cols = ["반", "번", "이름", "중간고사", "중간환산", "기말고사", "기말환산"]
        for j in range(n_perf):
            show_cols += [st.session_state.perf_names[j], f"{st.session_state.perf_names[j]}_환산"]
        show_cols += ["최종점수", "석차", "성취도", "석차등급", "비고"]
        show_cols = [c for c in show_cols if c in df.columns]

        st.dataframe(
            df[show_cols].style.format({
                "중간환산": "{:.2f}", "기말환산": "{:.2f}", "최종점수": "{:.2f}",
                **{f"{st.session_state.perf_names[j]}_환산": "{:.2f}" for j in range(n_perf)},
            }),
            use_container_width=True,
            hide_index=True,
        )

        # ── Excel 다운로드 ────────────────────────────────────
        @st.cache_data
        def to_excel(df):
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="성적산출결과")
            return buf.getvalue()

        st.download_button(
            "⬇️ 성적 결과 Excel 다운로드",
            data=to_excel(df[show_cols]),
            file_name=f"{st.session_state.grade}학년{st.session_state.semester}학기_{st.session_state.subject}_성적.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ═══════════════════════════════════════════════════════════════
# TAB 4 : 대시보드
# ═══════════════════════════════════════════════════════════════
with tab4:
    if not st.session_state.calculated or st.session_state.result_df is None:
        st.info("'점수 입력' 탭에서 점수를 입력하고 **성적 산출하기** 버튼을 눌러주세요.")
    else:
        df = st.session_state.result_df
        valid = df[df["최종점수"].notna()].copy()

        if len(valid) == 0:
            st.warning("유효한 점수 데이터가 없습니다.")
        else:
            # ── 1행: 점수 분포 히스토그램 + 성취도 파이 ──────────
            c1, c2 = st.columns(2)

            with c1:
                st.subheader("최종 점수 분포")
                fig_hist = px.histogram(
                    valid, x="최종점수", nbins=10,
                    color_discrete_sequence=["#4C72B0"],
                    labels={"최종점수": "점수", "count": "인원"},
                )
                fig_hist.update_layout(bargap=0.1, showlegend=False, height=320)
                st.plotly_chart(fig_hist, use_container_width=True)

            with c2:
                st.subheader("성취도 분포")
                grade_order = ["A", "B", "C", "D", "E"]
                grade_cnt = valid["성취도"].value_counts().reindex(grade_order, fill_value=0)
                fig_pie = px.pie(
                    names=grade_cnt.index,
                    values=grade_cnt.values,
                    color=grade_cnt.index,
                    color_discrete_map={"A":"#2ecc71","B":"#3498db","C":"#f39c12","D":"#e74c3c","E":"#95a5a6"},
                    hole=0.4,
                )
                fig_pie.update_layout(height=320)
                st.plotly_chart(fig_pie, use_container_width=True)

            # ── 2행: 석차 등급 막대 + 영역별 비교 ───────────────
            c3, c4 = st.columns(2)

            with c3:
                st.subheader("석차 등급 분포 (5등급)")
                grade_order = [1, 2, 3, 4, 5]
                grade_labels = ["1등급\n(상위10%)", "2등급\n(~34%)", "3등급\n(~66%)", "4등급\n(~90%)", "5등급\n(90%↓)"]
                numeric_ranks = valid["석차등급"].apply(lambda x: x if isinstance(x, int) else None).dropna()
                rank_cnt = numeric_ranks.value_counts().reindex(grade_order, fill_value=0)
                fig_bar = px.bar(
                    x=grade_labels,
                    y=rank_cnt.values,
                    labels={"x": "석차등급", "y": "인원"},
                    color=grade_labels,
                    color_discrete_sequence=["#2ECC71","#3498DB","#F39C12","#E67E22","#E74C3C"],
                    text_auto=True,
                )
                fig_bar.update_layout(showlegend=False, height=320)
                st.plotly_chart(fig_bar, use_container_width=True)

            with c4:
                st.subheader("영역별 평균 점수 (환산점수)")
                area_means = {}
                if "중간환산" in valid.columns:
                    area_means["중간고사"] = valid["중간환산"].mean()
                if "기말환산" in valid.columns:
                    area_means["기말고사"] = valid["기말환산"].mean()
                for j in range(st.session_state.perf_count):
                    col = f"{st.session_state.perf_names[j]}_환산"
                    if col in valid.columns:
                        area_means[st.session_state.perf_names[j]] = valid[col].mean()

                fig_area = px.bar(
                    x=list(area_means.keys()),
                    y=[round(v, 2) for v in area_means.values()],
                    labels={"x": "영역", "y": "평균 환산점수"},
                    color_discrete_sequence=["#1ABC9C"],
                    text_auto=True,
                )
                fig_area.update_layout(showlegend=False, height=320)
                st.plotly_chart(fig_area, use_container_width=True)

            # ── 3행: 상위 / 하위 학생 목록 ───────────────────────
            st.subheader("점수 상위 / 하위 학생")
            top_n = min(5, len(valid))
            c5, c6 = st.columns(2)
            with c5:
                st.markdown("**상위 5명**")
                st.dataframe(
                    valid.nsmallest(top_n, "석차")[["반","번","이름","최종점수","석차","성취도","석차등급"]],
                    hide_index=True, use_container_width=True,
                )
            with c6:
                st.markdown("**하위 5명**")
                st.dataframe(
                    valid.nlargest(top_n, "석차")[["반","번","이름","최종점수","석차","성취도","석차등급"]],
                    hide_index=True, use_container_width=True,
                )

            # ── 4행: 최소성취기준 미달 학생 ──────────────────────
            st.divider()
            min_threshold = st.number_input(
                "최소성취기준 기준점 (%)", min_value=0, max_value=100, value=40, step=5,
                help="최종점수가 이 비율 미만인 학생을 미달로 표시합니다."
            )
            below = df[df["최종점수"].notna() & (df["최종점수"] < min_threshold)]

            col_a, col_b = st.columns([1, 3])
            col_a.metric("⚠️ 최소성취기준 미달 학생", f"{len(below)}명",
                         delta=f"전체 {len(valid)}명 중", delta_color="off")

            if len(below) > 0:
                show_b = [c for c in ["반","번","이름","최종점수","석차","성취도","비고"] if c in below.columns]
                col_b.dataframe(
                    below[show_b].sort_values("최종점수"),
                    hide_index=True, use_container_width=True,
                )
            else:
                col_b.success(f"✅ 최종점수 {min_threshold}점 미만 학생 없음")
