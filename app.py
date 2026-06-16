# -*- coding: utf-8 -*-
import os, io, zipfile, tempfile, glob, datetime
import streamlit as st
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

def zip_dir(d):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(d):
            for fn in files:
                fp = os.path.join(root, fn); z.write(fp, os.path.relpath(fp, d))
    buf.seek(0); return buf

def all_company_files():
    out = []
    for loc in ["서울", "화성"]:
        for f in sorted(glob.glob(os.path.join(DATA, loc, "*.xlsx"))):
            n = os.path.basename(f)
            if n.startswith("_") or n.startswith("~$"): continue
            out.append((loc, f, n))
    return out

if not check_pw(): st.stop()
ADMIN = st.session_state.get("role") == "admin"
header()
if ADMIN:
    st.caption("매출·매입 세금계산서, 어음수취내역, 은행거래내역을 올리면 자동으로 종류를 판별하고 기존 자료에 신규만 추가합니다.")

with st.expander("🔎 거래처 검색 · 개별 다운로드", expanded=False):
    q = st.text_input("거래처명 검색", placeholder="예: 농협케미컬, 엘에스엠트론 …")
    files = all_company_files()
    if q:
        qn = q.replace(" ", "")
        hit = [(loc, f, n) for loc, f, n in files if qn in n.replace(" ", "")]
        if hit:
            st.caption(f"{len(hit)}건 일치")
            for loc, f, n in hit:
                c1, c2 = st.columns([4, 1]); c1.write(f"[{loc}] {n}")
                with open(f, "rb") as fh:
                    c2.download_button("다운로드", fh.read(), file_name=n, key="dl_" + f,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("일치하는 거래처가 없습니다.")
    else:
        st.caption(f"전체 {len(files)}개 거래처. 위 칸에 업체명을 입력하면 해당 파일만 받을 수 있어요.")

with st.expander("📥 현재 누적자료 전체 다운로드 (zip)", expanded=False):
    if os.path.isdir(DATA):
        st.download_button("전체 누적본 다운로드", zip_dir(DATA), file_name="누적자료_현재본.zip", mime="application/zip")

with st.expander("🛟 비상 복구 (직전본)", expanded=False):
    if os.path.exists(BACKUP_ZIP):
        st.caption("최근 갱신 직전 상태가 백업되어 있습니다. 문제가 생기면 한 번에 되돌릴 수 있어요.")
        with open(BACKUP_ZIP, "rb") as fh:
            st.download_button("직전본 다운로드 (zip)", fh.read(), file_name="누적자료_직전본.zip", mime="application/zip")
        if ADMIN and st.button("⏪ 직전본으로 복구"):
            if store.restore_previous(DATA, DATA_ZIP, BACKUP_ZIP):
                if GIT_TOKEN: store.git_commit_push(HERE, GIT_TOKEN, GIT_REPO, "restore previous baseline")
                st.success("직전본으로 복구했습니다."); st.rerun()
    else:
        st.caption("아직 직전본 백업이 없습니다. (첫 갱신 후 생성됩니다)")

if ADMIN:
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

        st.write("**파일 판별 결과**")
        for loc in uploads:
            d = res["detected"].get(loc, {})
            line = " / ".join(f"{k} {len(v)}건" for k, v in d.items() if v)
            if line: st.write(f"- {loc}: {line}")
        nc = res.get("new_companies", [])
        _stt = res["status"]
        st.markdown("**상태 요약**")
        _cards = [("신규 업체", f"{len(nc)}곳", True)] + [(_k, str(_stt.get(_k, 0)), False) for _k in ["완납", "진행", "미수", "장기미수"]]
        _h = "<div style='display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin:4px 0 14px;'>"
        for _lab, _val, _hi in _cards:
            _bg = NAVY if _hi else "#EAF1FB"; _lc = "#ccdcf5" if _hi else "#3A5C8A"; _vc = "#ffffff" if _hi else NAVY
            _h += f"<div style='background:{_bg};border-radius:10px;padding:12px 14px;'><div style='font-size:12px;color:{_lc};'>{_lab}</div><div style='font-size:24px;font-weight:700;color:{_vc};'>{_val}</div></div>"
        _h += "</div>"
        st.markdown(_h, unsafe_allow_html=True)
        if nc: st.caption("신규: " + ", ".join(nc[:20]) + (" 외" if len(nc) > 20 else ""))
        if res["unassigned"]:
            st.warning(f"미배정 입금 {len(res['unassigned'])}건 (입금자명 매칭 실패 — 별칭표 보완 필요)")

        if not ok:
            st.error("검증 실패 — 누적본을 갱신하지 않았습니다. 직전본은 그대로 유지됩니다.")
            st.write(probs[:15])
            st.download_button("⚠ 검증 전 결과 받아보기 (zip)", zip_dir(out_dir),
                               file_name="누적자료_검증전.zip", mime="application/zip")
        else:
            store.apply_update(out_dir, DATA, DATA_ZIP, BACKUP_ZIP)
            open(BASIS, "w", encoding="utf-8").write(datetime.date.today().isoformat())
            saved, msg = store.git_commit_push(HERE, GIT_TOKEN, GIT_REPO, "update baseline via app")
            st.success("검증 통과 — 누적본을 갱신했고 직전본을 백업했습니다." +
                       ("  (GitHub 영구저장 완료)" if saved else f"  (GitHub 미저장: {msg})"))
            st.download_button("📥 갱신된 누적본 다운로드 (zip)", zip_dir(DATA),
                               file_name="누적자료_최신본.zip", mime="application/zip", type="primary")
