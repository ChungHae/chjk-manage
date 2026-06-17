# -*- coding: utf-8 -*-
import os, io, zipfile, tempfile, glob, datetime, re
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import pipeline, store

st.set_page_config(page_title="충해전기 관리시스템", page_icon="📒", layout="wide")
st.markdown("<style>.block-container{max-width:1180px;margin:0 auto;padding-top:2.2rem;padding-left:3rem;padding-right:3rem;}</style>", unsafe_allow_html=True)
HERE = os.path.dirname(os.path.abspath(__file__))
DATA_ZIP = os.path.join(HERE, "data.zip")
BACKUP_ZIP = os.path.join(HERE, "_직전본.zip")
BASIS = os.path.join(HERE, "_기준일.txt")
DATA = store.ensure_data(DATA_ZIP, os.path.join(tempfile.gettempdir(), "misu_work"))
LOGO = "https://chjk.co.kr/web/upload/category/logo/v2_c8bcd54017bc5f8880bb32d3de5333e6_BWrlZyel0_top.jpg"
NAVY = "#1B3A6B"
GIT_TOKEN = st.secrets.get("github_token", os.environ.get("GITHUB_TOKEN", ""))
GIT_REPO = st.secrets.get("github_repo", os.environ.get("GITHUB_REPO", ""))

def basis_date():
    try: return datetime.date.fromisoformat(open(BASIS, encoding="utf-8").read().strip())
    except Exception: return None

def header(full=True):
    title = "미수관리 시스템" if full else "관리 시스템"
    bd = basis_date()
    sub = f"&nbsp;&nbsp;·&nbsp;&nbsp;현재 자료: {bd.year}년 {bd.month}월 {bd.day}일 기준" if (full and bd) else ""
    st.markdown(
        f"<div style='display:flex;align-items:center;gap:18px;border-bottom:2px solid {NAVY};padding-bottom:12px;margin:4px 0 16px;'>"
        f"<img src='{LOGO}' style='height:52px;width:auto;object-fit:contain;' onerror=\"this.style.display='none'\">"
        f"<div><div style='font-size:22px;font-weight:600;color:{NAVY};line-height:1.25;'>{title}</div>"
        f"<div style='font-size:12px;letter-spacing:2px;color:#888;'>PNEUMATIC MALL{sub}</div></div></div>",
        unsafe_allow_html=True)

def check_pw():
    if st.session_state.get("auth"): return True
    header(full=False)
    pw_admin = st.secrets.get("password_admin", st.secrets.get("password", os.environ.get("APP_PW", "chunghae")))
    pw_view = st.secrets.get("password_view", os.environ.get("APP_PW_VIEW", ""))
    with st.form("login_form"):
        x = st.text_input("비밀번호", type="password")
        submitted = st.form_submit_button("로그인")
    if submitted:
        if x == pw_admin: st.session_state.update(auth=True, role="admin"); st.rerun()
        elif pw_view and x == pw_view: st.session_state.update(auth=True, role="view"); st.rerun()
        else: st.error("비밀번호가 틀립니다.")
    return False

def zip_bytes(d, dirs_only=None):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(d):
            for fn in files:
                fp = os.path.join(root, fn); rel = os.path.relpath(fp, d).replace("\\", "/")
                if dirs_only and not any(rel.startswith(p + "/") for p in dirs_only): continue
                z.write(fp, rel)
    return buf.getvalue()

def zip_backup_customers():
    """직전본.zip에서 서울/화성 거래처 파일만 추려 재압축(지원표 제외)."""
    buf = io.BytesIO()
    with zipfile.ZipFile(BACKUP_ZIP) as src, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in src.namelist():
            nn = n.replace("\\", "/")
            if nn.startswith("서울/") or nn.startswith("화성/"): z.writestr(n, src.read(n))
    return buf.getvalue()

def all_company_files():
    out = []
    for loc in ["서울", "화성"]:
        for f in sorted(glob.glob(os.path.join(DATA, loc, "*.xlsx"))):
            n = os.path.basename(f)
            if n.startswith("_") or n.startswith("~$"): continue
            out.append((loc, f, n))
    return out

def render_result(R):
    st.divider()
    st.write("**파일 판별 결과**")
    for loc, d in R["detected"].items():
        line = " / ".join(f"{k} {len(v)}건" for k, v in d.items() if v)
        if line: st.write(f"- {loc}: {line}")
    # 상태 카드
    nc = R["new_companies"]; stt = R["status"]
    st.markdown("**상태 요약**")
    cards = [("신규 업체", f"{len(nc)}곳", True)] + [(k, str(stt.get(k, 0)), False) for k in ["완납", "진행", "미수", "장기미수"]]
    h = "<div style='display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin:4px 0 14px;'>"
    for lab, val, hi in cards:
        bg = NAVY if hi else "#EAF1FB"; lc = "#ccdcf5" if hi else "#3A5C8A"; vc = "#ffffff" if hi else NAVY
        h += f"<div style='background:{bg};border-radius:10px;padding:12px 14px;'><div style='font-size:12px;color:{lc};'>{lab}</div><div style='font-size:24px;font-weight:700;color:{vc};'>{val}</div></div>"
    st.markdown(h + "</div>", unsafe_allow_html=True)
    # 이번에 갱신·신규 처리된 거래처 (접기/펼치기)
    chg = [{"구분": s[8], "지역": s[0], "거래처": s[1], "상태": s[3],
            "계산서 추가": int(s[6]), "입금 추가": int(s[7])}
           for s in R["summary"] if (s[6] or s[7])]
    if chg:
        with st.expander(f"이번에 갱신·신규 처리된 거래처 {len(chg)}곳 자세히 보기 / 내려받기"):
            cdf = pd.DataFrame(chg).sort_values(["구분", "지역", "거래처"]).reset_index(drop=True)
            st.dataframe(cdf, use_container_width=True, hide_index=True)
            st.download_button("이 목록 CSV 다운로드", cdf.to_csv(index=False).encode("utf-8-sig"),
                               file_name="갱신_신규_거래처.csv", mime="text/csv", key="dl_chg")
    else:
        st.caption("이번에 갱신·신규 처리된 거래처가 없습니다. (동일 자료 재처리 등)")
    # 미배정
    if R["unassigned"]:
        st.warning(f"미배정 입금 {len(R['unassigned'])}건 (입금자명 매칭 실패 — 별칭표 보완 필요)")
        with st.expander(f"미배정 입금 {len(R['unassigned'])}건 자세히 보기 / 내려받기"):
            udf = pd.DataFrame(R["unassigned"], columns=["지역", "날짜", "금액", "입금자명", "유형"])
            udf = udf.sort_values(["지역", "입금자명", "날짜"]).reset_index(drop=True)
            st.dataframe(udf, use_container_width=True, hide_index=True)
            st.download_button("미배정 목록 CSV 다운로드", udf.to_csv(index=False).encode("utf-8-sig"),
                               file_name="미배정입금.csv", mime="text/csv", key="dl_un")
    # 최종 결과 + 다운로드
    if R["ok"]:
        st.success("검증 통과 — 누적본을 갱신했고 직전본을 백업했습니다." + R["saved_note"])
        st.download_button("📥 갱신된 누적본 다운로드 (zip)", R["zip_bytes"],
                           file_name="누적자료_최신본.zip", mime="application/zip", type="primary", key="dl_final")
    else:
        st.error("검증 실패 — 누적본을 갱신하지 않았습니다. 직전본은 그대로 유지됩니다.")
        st.write(R["probs"][:15])
        st.download_button("⚠ 검증 전 결과 받아보기 (zip)", R["zip_bytes"],
                           file_name="누적자료_검증전.zip", mime="application/zip", key="dl_final")

if not check_pw(): st.stop()
ADMIN = st.session_state.get("role") == "admin"
header()
if ADMIN:
    st.markdown(f"<div style='margin:-6px 0 8px;'><span style='background:{NAVY};color:#fff;font-size:12px;font-weight:600;padding:3px 12px;border-radius:6px;'>🔑 관리자 계정</span></div>", unsafe_allow_html=True)
    st.caption("매출·매입 세금계산서, 어음수취내역, 은행거래내역을 올리면 자동으로 종류를 판별하고 기존 자료에 신규만 추가합니다.")

def _cust_label(loc, n):
    m = re.search(r"\((\d{3}-\d{2}-\d{5})\)", n)
    biz = m.group(1) if m else ""
    nm = re.sub(r"^[^)]*\)\s*", "", n).rsplit(" (", 1)[0]
    return f"[{loc}] {nm}" + (f" ({biz})" if biz else "")

with st.expander("🔎 거래처 검색 · 개별 다운로드", expanded=False):
    files = all_company_files()
    _label_map = {}
    for loc, f, n in files:
        _label_map[_cust_label(loc, n)] = (f, n)
    _opts = list(_label_map.keys())
    sel = st.selectbox(f"거래처명 일부만 입력하면 자동으로 후보가 나옵니다 (전체 {len(files)}곳) · 엔터(또는 클릭) 후 바로 입력하면 새로 검색됩니다",
                       _opts, index=None, placeholder="거래처명 입력 (예: 농협)", key="cust_sel")
    if sel and sel in _label_map:
        f, n = _label_map[sel]
        with open(f, "rb") as fh:
            st.download_button(f"📥 {sel} 다운로드", fh.read(), file_name=n, key="dl_sel",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with st.expander("📥 현재 누적자료 전체 다운로드 (zip)", expanded=False):
    if os.path.isdir(DATA):
        st.download_button("전체 누적본 다운로드 (서울·화성 거래처)", zip_bytes(DATA, ["서울", "화성"]), file_name="누적자료_현재본.zip", mime="application/zip", key="dl_all")

with st.expander("🛟 비상 복구 (직전본)", expanded=False):
    if os.path.exists(BACKUP_ZIP):
        st.caption("최근 갱신 직전 상태가 백업되어 있습니다. 문제가 생기면 한 번에 되돌릴 수 있어요.")
        bc1, bc2 = st.columns(2)
        bc1.download_button("직전본 다운로드 (서울·화성)", zip_backup_customers(), file_name="누적자료_직전본.zip", mime="application/zip", key="dl_bak")
        if ADMIN and bc2.button("⏪ 직전본으로 복구"):
            if store.restore_previous(DATA, DATA_ZIP, BACKUP_ZIP):
                if GIT_TOKEN: store.git_commit_push(HERE, GIT_TOKEN, GIT_REPO, "restore previous baseline")
                st.session_state.pop("result", None)
                st.success("직전본으로 복구했습니다."); st.rerun()
    else:
        st.caption("아직 직전본 백업이 없습니다. (첫 갱신 후 생성됩니다)")

if ADMIN:
    with st.expander("✏️ 거래처 파일 직접 수정 후 업로드 (수동 교체)", expanded=False):
        st.caption("거래처 파일을 받아 직접 고친 뒤 여기 올리면 그 거래처를 교체합니다. 사업자번호로 자동 인식하며, 교체 전 직전본이 백업됩니다.")
        m_ups = st.file_uploader("수정한 거래처 파일(.xlsx) 업로드", accept_multiple_files=True, type=["xlsx"], key="manual_edit")
        if m_ups:
            parsed = []
            for uf in m_ups:
                nm = uf.name; raw = bytes(uf.getbuffer())
                mb = re.search(r"\((\d{3}-\d{2}-\d{5})\)", nm)
                ml = re.search(r"(서울|화성)\)", nm)
                loc = ml.group(1) if ml else ("화성" if "화성" in nm else "서울")
                try: load_workbook(io.BytesIO(raw)); ok = True
                except Exception: ok = False
                parsed.append((loc, nm, (mb.group(1) if mb else None), ok, raw))
            for loc, nm, biz, ok, _ in parsed:
                st.write(f"- [{loc}] {nm} → " + ("✅ 교체 준비됨" if (ok and biz) else "⚠️ 사업자번호/형식 확인 필요(건너뜀)"))
            good = [(loc, nm, b) for loc, nm, biz, ok, b in parsed if ok and biz]
            if good and st.button(f"{len(good)}개 거래처 교체 (직전본 백업 후)", type="primary", key="manual_apply"):
                store.replace_customer_files(DATA, good, DATA_ZIP, BACKUP_ZIP)
                saved, msg = store.git_commit_push(HERE, GIT_TOKEN, GIT_REPO, "manual edit replace")
                st.session_state.pop("result", None)
                st.success(f"{len(good)}개 거래처 교체 완료." + ("  (GitHub 영구저장)" if saved else f"  (미저장: {msg})"))
                st.rerun()

    st.subheader("자료 업로드")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**서울** (신한·기업·하나·농협·우리 등)")
        up_seoul = st.file_uploader("서울 자료", accept_multiple_files=True, key="seoul",
                                    type=["xls", "xlsx"], label_visibility="collapsed")
    with col2:
        st.markdown("**화성** (신한)")
        up_hwa = st.file_uploader("화성 자료", accept_multiple_files=True, key="hwa",
                                  type=["xls", "xlsx"], label_visibility="collapsed")

    if st.button("🚀 처리 시작", type="primary", disabled=not (up_seoul or up_hwa)):
        work = tempfile.mkdtemp(); up_dir = os.path.join(work, "up"); out_dir = os.path.join(work, "out")
        os.makedirs(up_dir, exist_ok=True); os.makedirs(out_dir, exist_ok=True)
        uploads = {"서울": [], "화성": []}
        for loc, ulist in [("서울", up_seoul or []), ("화성", up_hwa or [])]:
            for f in ulist:
                p = os.path.join(up_dir, f"{loc}_{f.name}"); open(p, "wb").write(f.getbuffer()); uploads[loc].append(p)
        log_area = st.empty(); logs = []
        def prog(m): logs.append(m); log_area.code("\n".join(logs[-12:]))
        with st.spinner("처리 중… (파일이 많으면 수십 초 걸릴 수 있어요)"):
            try:
                res = pipeline.process(uploads, DATA if os.path.isdir(DATA) else up_dir, out_dir, progress=prog)
            except Exception as e:
                st.error(f"처리 오류: {e}"); st.stop()
            ok, probs = pipeline.verify(out_dir, DATA)
            if ok:
                store.apply_update(out_dir, DATA, DATA_ZIP, BACKUP_ZIP)
                open(BASIS, "w", encoding="utf-8").write(res.get("basis") or datetime.date.today().isoformat())
                saved, msg = store.git_commit_push(HERE, GIT_TOKEN, GIT_REPO, "update baseline via app")
                note = "  (GitHub 영구저장 완료)" if saved else f"  (GitHub 미저장: {msg})"
                zb = zip_bytes(DATA)
            else:
                note = ""; zb = zip_bytes(out_dir)
        st.session_state["result"] = {
            "detected": res["detected"], "status": res["status"], "new_companies": res.get("new_companies", []),
            "summary": res["summary"], "unassigned": res["unassigned"],
            "ok": ok, "probs": probs, "zip_bytes": zb, "saved_note": note,
        }
        log_area.empty()

    if st.session_state.get("result"):
        render_result(st.session_state["result"])
