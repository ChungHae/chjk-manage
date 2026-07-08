# -*- coding: utf-8 -*-
import os, io, zipfile, tempfile, glob, datetime, re
import hashlib, requests
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import pipeline, store

st.set_page_config(page_title="충해전기 관리시스템", page_icon="📒", layout="wide")
import streamlit.components.v1 as _components
_components.html("<script>window.parent.document.title='충해전기 관리시스템';</script>", height=0)
st.markdown("<style>[data-testid='stMain']{scrollbar-gutter:stable;}.block-container{max-width:1320px;margin:0 auto;padding-top:1.4rem;padding-left:3rem;padding-right:3rem;}@media(max-width:640px){.block-container{padding-left:0.7rem!important;padding-right:0.7rem!important;}}</style><style>iframe[height='0']{display:none;}.st-key-nytrig{position:fixed!important;left:-9999px!important;top:0!important;width:1px!important;height:1px!important;overflow:hidden!important;}a[href$='page_nyung'],[data-testid='stTopNavLink'][href$='page_nyung'],[data-testid='stTopNavLinkContainer']:has(a[href$='page_nyung']),li:has(a[href$='page_nyung']),[data-testid='stSidebarNavItems'] li:has(a[href$='page_nyung']){display:none!important;}</style>", unsafe_allow_html=True)
HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = "https://chjk.co.kr/web/upload/category/logo/v2_c8bcd54017bc5f8880bb32d3de5333e6_BWrlZyel0_top.jpg"
NAVY = "#1B3A6B"
GIT_TOKEN = st.secrets.get("github_token", os.environ.get("GITHUB_TOKEN", ""))
DATA_REPO = st.secrets.get("github_data_repo", st.secrets.get("github_repo", os.environ.get("GITHUB_REPO", "")))

@st.cache_resource(show_spinner=False)
def _data_repo_dir():
    return store.clone_data_repo(DATA_REPO, GIT_TOKEN, os.path.join(tempfile.gettempdir(), "chjk_data_repo"))

_REPO_DIR = _data_repo_dir() or HERE   # 자료 저장소(비공개). 실패 시 코드폴더(번들 data.zip) 대체
DATA_OK = bool(_data_repo_dir()) or os.path.exists(os.path.join(HERE, "data.zip"))
DATA_ZIP = os.path.join(_REPO_DIR, "data.zip")
BACKUP_ZIP = os.path.join(_REPO_DIR, "_직전본.zip")
BASIS = os.path.join(_REPO_DIR, "_기준일.txt")
DATA = store.ensure_data(DATA_ZIP, os.path.join(tempfile.gettempdir(), "misu_work"))

def basis_date():
    try: return datetime.date.fromisoformat(open(BASIS, encoding="utf-8").read().strip())
    except Exception: return None

def header(full=True, title=None):
    title = title or ("미수관리 시스템" if full else "관리 시스템")
    bd = basis_date()
    sub = f"현재 자료: {bd.year}년 {bd.month}월 {bd.day}일 기준" if (full and bd) else ""
    st.markdown(
        f"<div style='display:flex;align-items:center;gap:18px;border-bottom:2px solid {NAVY};padding-bottom:12px;margin:6px 0 16px;'>"
        f"<img src='{LOGO}' style='height:52px;width:auto;object-fit:contain;' onerror=\"this.style.display='none'\">"
        f"<div><div style='font-size:22px;font-weight:600;color:{NAVY};line-height:1.25;'>{title}</div>"
        f"<div style='font-size:12px;letter-spacing:2px;color:#888;'>{sub}</div></div></div>",
        unsafe_allow_html=True)

# ── 업무관리 앱과 로그인 공유 (동일 Firebase 계정) ──
_FB_API_KEY = "AIzaSyBTYey9GzRIRRqiiZoQ3gpxI-Ty1BhXyZU"
_FB_DB_URL  = "https://chjk-scheduler-default-rtdb.asia-southeast1.firebasedatabase.app"
_FB_PATH    = "teamdata_test"
MISU_ADMIN_IDS = ["chjk", "김소준"]   # 자료처리 전체 권한 (그 외 계정은 조회 전용)

def _sha256(v):
    return hashlib.sha256(str(v).encode("utf-8")).hexdigest()

def _synth_email(uid):
    return "u" + _sha256(uid)[:32] + "@chjk-scheduler.web.app"

def _fb_login(uid, pw):
    """업무관리와 동일한 Firebase 계정으로 검증. return (ok, role, err)."""
    if not uid or not pw:
        return False, None, "성명과 비밀번호를 입력하세요."
    try:
        r = requests.post(
            "https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=" + _FB_API_KEY,
            json={"email": _synth_email(uid), "password": _sha256(pw), "returnSecureToken": True},
            timeout=15)
    except Exception:
        return False, None, "서버 연결에 실패했습니다. 잠시 후 다시 시도하세요."
    if r.status_code != 200:
        return False, None, "성명 또는 비밀번호가 올바르지 않습니다."
    d = r.json(); token = d.get("idToken"); fuid = d.get("localId")
    try:
        a = requests.get(_FB_DB_URL + "/" + _FB_PATH + "_authorized/" + str(fuid) + ".json?auth=" + str(token), timeout=15)
        allowed = (a.status_code == 200 and a.json() is True)
    except Exception:
        allowed = False
    if not allowed:
        return False, None, "접근이 허용되지 않은 계정입니다. 관리자에게 문의하세요."
    import time as _t
    st.session_state["_fb_tok"] = token
    st.session_state["_fb_ref"] = d.get("refreshToken")
    st.session_state["_fb_tok_at"] = _t.time()
    role = "admin" if uid in MISU_ADMIN_IDS else "view"
    return True, role, ""

def _fb_id_token():
    """저장된 토큰 재사용, 만료 임박 시 refresh_token으로 갱신. 실패 시 None."""
    import time as _t
    tok = st.session_state.get("_fb_tok"); at = st.session_state.get("_fb_tok_at", 0)
    if tok and (_t.time() - at) < 3300:
        return tok
    ref = st.session_state.get("_fb_ref")
    if not ref: return None
    try:
        r = requests.post("https://securetoken.googleapis.com/v1/token?key=" + _FB_API_KEY,
                          data={"grant_type": "refresh_token", "refresh_token": ref}, timeout=15)
        if r.status_code == 200:
            j = r.json(); nt = j.get("id_token")
            st.session_state["_fb_tok"] = nt
            st.session_state["_fb_ref"] = j.get("refresh_token", ref)
            st.session_state["_fb_tok_at"] = _t.time()
            return nt
    except Exception:
        return None
    return None

def _publish_misu_summary(summary):
    """미수 요약을 Firebase 전용 키에 저장(업무관리 카드용). 실패해도 화면에 영향 없음."""
    try:
        import json as _json
        sig = _json.dumps({k: v for k, v in summary.items() if k != "updatedAt"},
                          ensure_ascii=False, sort_keys=True)
        if st.session_state.get("_misu_pub_sig") == sig: return
        tok = _fb_id_token()
        if not tok: return
        r = requests.put(_FB_DB_URL + "/" + _FB_PATH + "_misu_summary.json?auth=" + tok,
                         json=summary, timeout=10)
        if r.status_code == 200:
            st.session_state["_misu_pub_sig"] = sig
    except Exception:
        pass

_LOGIN_CSS = """
<style>
[data-testid='stSidebar'],[data-testid='stSidebarCollapsedControl'],[data-testid='stSidebarCollapseButton']{display:none!important;}
[data-testid='stMainBlockContainer'],.block-container{max-width:460px!important;padding-top:7vh!important;}
[data-testid='stForm']{border:1px solid rgba(26,26,26,.12);border-radius:14px;box-shadow:0 10px 34px rgba(27,58,107,.10);background:#fff;padding:0 22px 22px!important;}
[data-testid='stForm'] label{font-size:13px!important;font-weight:600!important;color:#374151!important;}
[data-testid='stForm'] input{background:#f7f9fc!important;border:1px solid rgba(26,26,26,.2)!important;border-radius:8px!important;padding:11px 12px!important;font-size:14px!important;}
[data-testid='stForm'] input:focus{border-color:#1B3A6B!important;background:#fff!important;box-shadow:none!important;}
[data-testid='stFormSubmitButton'] button{width:100%!important;background:#1B3A6B!important;border:none!important;border-radius:8px!important;padding:12px!important;}
[data-testid='stFormSubmitButton'] button:hover{background:#14305c!important;}
[data-testid='stFormSubmitButton'] button p{color:#fff!important;font-weight:700!important;font-size:14px!important;}
.misu-auth-head{display:flex;align-items:center;gap:12px;padding:18px 22px;border-bottom:2px solid #1B3A6B;margin:0 -22px 18px;}
.misu-auth-head img{height:34px;width:auto;display:block;}
.misu-auth-head span{font-size:25px;line-height:34px;font-weight:600;color:#1B3A6B;letter-spacing:-.3px;}
.misu-auth-foot{text-align:center;font-size:11px;color:#9ca3af;margin-top:14px;}
</style>
"""

def check_pw():
    if st.session_state.get("auth"): return True
    box = st.empty()
    with box.container():
        st.markdown(_LOGIN_CSS, unsafe_allow_html=True)
        with st.form("login_form"):
            st.markdown("<div class='misu-auth-head'><img src='" + LOGO + "' alt='충해전기(주)' onerror=\"this.style.display='none'\"><span>미수 관리</span></div>", unsafe_allow_html=True)
            uid = st.text_input("성명", placeholder="성명")
            pw = st.text_input("비밀번호", type="password", placeholder="비밀번호")
            submitted = st.form_submit_button("로그인")
            st.markdown("<div class='misu-auth-foot'>충해전기(주) 내부 운영 시스템 · 관계자 외 접근 금지</div>", unsafe_allow_html=True)
    if submitted:
        ok, role, err = _fb_login((uid or "").strip(), pw or "")
        if ok:
            box.empty()
            st.session_state.update(auth=True, role=role, uid=(uid or "").strip())
            st.rerun()
        else:
            st.error(err)
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

def _cust_label(loc, n):
    m = re.search(r"\((\d{3}-\d{2}-\d{5})\)", n)
    biz = m.group(1) if m else ""
    nm = re.sub(r"^[^)]*\)\s*", "", n).rsplit(" (", 1)[0]
    return f"[{loc}] {nm}" + (f" ({biz})" if biz else "")

def _duerules():
    duer = {}
    try:
        wb = load_workbook(os.path.join(DATA, "_거래처기준설정표.xlsx"), data_only=True)
        ws = wb["기준설정"] if "기준설정" in wb.sheetnames else wb.active
        for r in range(2, ws.max_row + 1):
            b = ws.cell(r, 3).value; ru = ws.cell(r, 4).value
            if b and ru: duer[str(b).strip()] = str(ru).strip()
        wb.close()
    except Exception:
        pass
    return duer

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

# ===== 미수 현황 대시보드 페이지 =====
@st.cache_data(show_spinner="미수 현황 계산 중…")
def _dashboard_html(_fp, basis_iso, admin):
    import dashboard as _dash
    return _dash.render(DATA, datetime.date.fromisoformat(basis_iso), _duerules(), admin=admin)

@st.cache_data(show_spinner=False)
def _longoverdue_cached(_fp):
    return _longoverdue_list()

@st.cache_resource(show_spinner=False)
def _nyung_book(_fp):
    import nyung
    return nyung.load_book(_REPO_DIR)

def _data_fp():
    return (os.path.getmtime(DATA_ZIP) if os.path.exists(DATA_ZIP) else 0,
            len(glob.glob(os.path.join(DATA, "*", "*.xlsx"))))

def _tpl_mtime(name):
    """템플릿(HTML) 수정시각 — 캐시 키에 넣어, 템플릿만 바꿔도(재배포 시) 화면이 새로 렌더되게 함."""
    try:
        return os.path.getmtime(os.path.join(HERE, name))
    except Exception:
        return 0

@st.cache_data(show_spinner=False)
def _sales_summary_data(_fp):
    try:
        import sales as _sales
        bd = _sales.build_data(DATA, _REPO_DIR)
        if not bd: return {}
        years = bd.get("years") or {}; summ = bd.get("summary") or {}
        if not years: return {}
        cy = max(years.keys())
        sby = summ.get("매출 (계산서 기준)", {}); pby = summ.get("매출총이익", {}); mby = summ.get("마진율", {})
        months = [e for e in years.get(cy, []) if e.get("매출합")]
        lm = months[-1] if months else None
        return {
            "sales_year": cy,
            "sales_year_total": int(sby.get(cy, 0) or 0),
            "sales_year_profit": int(pby.get(cy, 0) or 0),
            "sales_year_margin": round(float(mby.get(cy, 0) or 0) * 100, 1),
            "sales_latest_month": (lm.get("월") if lm else ""),
            "sales_latest_month_total": int(lm.get("매출합", 0)) if lm else 0,
        }
    except Exception:
        return {}

@st.cache_data(show_spinner=False)
def _misu_summary_data(_fp, basis_iso):
    import dashboard as _dash
    bd = datetime.date.fromisoformat(basis_iso)
    data = _dash.compute_data(DATA, bd, _duerules())
    unpaid = [o for o in data if o["status"] in ("미수", "장기미수")]
    longx = [o for o in data if o["status"] == "장기미수"]
    return {
        "basis": basis_iso,
        "updatedAt": datetime.datetime.now().isoformat(timespec="seconds"),
        "customer_count": len(data),
        "unpaid_count": len(unpaid),
        "longoverdue_count": len(longx),
        "total_unpaid": int(sum(o["amt"] for o in unpaid)),
        "seoul_unpaid": int(sum(o["amt"] for o in unpaid if o["reg"] == "서울")),
        "hwaseong_unpaid": int(sum(o["amt"] for o in unpaid if o["reg"] == "화성")),
        "longoverdue_amt": int(sum(o["amt"] for o in longx)),
    }

def page_dashboard():
    bd = basis_date() or datetime.date.today()
    if not os.path.isdir(DATA):
        header(); st.error("자료를 불러올 수 없습니다."); return
    try:
        _fp = _data_fp() + (_tpl_mtime("dashboard_template.html"),)
        _components.html(_dashboard_html(_fp, bd.isoformat(), ADMIN), height=760, scrolling=True)
        # 팝업의 '내용증명 작성' 버튼이 같은 출처(same-origin)로 클릭할 숨은 트리거(거래처별). 관리자·실무자 모두.
        with st.container(key="nytrig"):
            for _o in _longoverdue_cached(_fp):
                if st.button("NYO_" + _o["biz"], key="nyo_" + _o["biz"]):
                    st.session_state["_routed_biz"] = _o["biz"]
                    st.session_state["_ny_applied"] = None
                    st.switch_page(_nyung_page)
    except Exception as e:
        header(); st.error(f"대시보드 생성 오류: {e}")
    try:
        _sum = dict(_misu_summary_data(_fp, bd.isoformat()))
        try: _sum.update(_sales_summary_data(_fp))
        except Exception: pass
        _publish_misu_summary(_sum)
    except Exception:
        pass

# ===== 매출 현황 페이지 =====
@st.cache_data(show_spinner="매출 현황 불러오는 중…")
def _sales_html(_fp, basis_iso, admin):
    import sales as _sales
    return _sales.render(DATA, admin=admin, basis_iso=basis_iso)

def page_sales():
    if not os.path.isdir(DATA):
        header(title="매출 현황"); st.error("자료를 불러올 수 없습니다."); return
    if not os.path.exists(os.path.join(DATA, "_매출자료.json")):
        header(title="매출 현황")
        st.warning("매출 자료(_매출자료.json)가 없습니다. chjk-data의 data.zip에 매출 자료를 포함해 주세요.")
        return
    try:
        _bd = basis_date()
        _fp = (os.path.getmtime(DATA_ZIP) if os.path.exists(DATA_ZIP) else 0, _tpl_mtime("sales_template.html"))
        _components.html(_sales_html(_fp, _bd.isoformat() if _bd else "", ADMIN), height=900, scrolling=True)
    except Exception as e:
        header(title="매출 현황"); st.error(f"매출 현황 생성 오류: {e}")

# ===== 내용증명 페이지 =====
def _longoverdue_list():
    """현재 장기미수 거래처(실시간 계산). 상태가 바뀌면 자동 반영."""
    import dashboard as _dash
    bd = basis_date() or datetime.date.today()
    data = _dash.compute_data(DATA, bd, _duerules())
    lt = [o for o in data if o.get("status") == "장기미수"]
    lt.sort(key=lambda o: (o["reg"], -o["amt"]))
    return lt

def _fontconfig_file():
    """맑은 고딕/Malgun Gothic → Noto Sans CJK KR 로 강제 매핑(서버 PDF가 워드와 동일하게 보이도록)."""
    conf = os.path.join(tempfile.gettempdir(), "chjk_fonts.conf")
    try:
        if not os.path.exists(conf):
            with open(conf, "w", encoding="utf-8") as f:
                f.write('<?xml version="1.0"?>\n<!DOCTYPE fontconfig SYSTEM "fonts.dtd">\n<fontconfig>\n'
                        '  <include ignore_missing="yes">/etc/fonts/fonts.conf</include>\n'
                        '  <match target="pattern"><test name="family"><string>맑은 고딕</string></test>'
                        '<edit name="family" mode="assign" binding="strong"><string>Noto Sans CJK KR</string></edit></match>\n'
                        '  <match target="pattern"><test name="family"><string>Malgun Gothic</string></test>'
                        '<edit name="family" mode="assign" binding="strong"><string>Noto Sans CJK KR</string></edit></match>\n'
                        '</fontconfig>\n')
        return conf
    except Exception:
        return None

def _docx_to_pdf(docx_path, out_dir):
    """LibreOffice headless 로 docx→pdf. 폰트 별칭(맑은고딕→Noto)으로 워드와 동일하게 렌더. 실패 시 None."""
    import subprocess
    base = os.path.splitext(os.path.basename(docx_path))[0]
    env = dict(os.environ)
    conf = _fontconfig_file()
    if conf:
        env["FONTCONFIG_FILE"] = conf
    lohome = os.path.join(tempfile.gettempdir(), "lohome")
    try:
        os.makedirs(lohome, exist_ok=True); env.setdefault("HOME", lohome)
    except Exception:
        pass
    for soffice in ("libreoffice", "soffice"):
        try:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
                           check=True, timeout=120, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)
            pdf = os.path.join(out_dir, base + ".pdf")
            if os.path.exists(pdf):
                return pdf
        except Exception:
            continue
    return None

def _pdf_pages(path):
    try:
        from pypdf import PdfReader
        return len(PdfReader(path).pages)
    except Exception:
        return 1

def _reformat_amt(_k):
    """미수액 입력값을 천단위 콤마로 자동 정리(on_change 콜백 — 위젯 키 수정 허용 구간)."""
    d = re.sub(r"[^0-9]", "", str(st.session_state.get(_k, "")))
    st.session_state[_k] = f"{int(d):,}" if d else ""

def _make_pdf_fit(data, pdf_src, workdir):
    """PDF는 노토 폰트로 생성. 항목간격을 키워 하단까지 채우되, 페이지수를 보며 1페이지를 유지하도록 자동 조절.
    서버 렌더링이 sandbox보다 촘촘해도 알아서 더 채운다. 성공값은 세션에 캐시해 다음부터 빠르게."""
    import nyung
    cached = st.session_state.get("_pdf_sp")
    cands = ([cached] if cached else []) + [9, 7, 6, 5, 4, 3]
    seen = set(); order = [c for c in cands if not (c in seen or seen.add(c))]
    last = None
    for sp in order:
        # 본문 번호 사이(1-2,2-3,3-4,4-5,6-7)와 구간 사이(수신자↔발신자, 발신자↔촉구의 건, 입금표 위·아래, 실선↔날짜, 회사명↔명판)를 균일하게 띄워 고르게 채움
        nyung.make_docx(data, _REPO_DIR, pdf_src, font="Noto Sans CJK KR", date_before=2, item_after=sp, spread=sp)
        cand = _docx_to_pdf(pdf_src, workdir)
        if not cand:
            return None
        last = cand
        if _pdf_pages(cand) == 1:
            st.session_state["_pdf_sp"] = sp
            return cand
    return last

def page_nyung():
    import nyung
    header(title="내용증명")
    if st.button("← 미수 현황으로 돌아가기"):
        st.switch_page(_dashboard_page)
    if not os.path.isdir(DATA):
        st.error("자료를 불러올 수 없습니다."); return
    miss = [f for f in ("stamp_guro.png", "stamp_hwaseong.png", "인감_보정_투명.png")
            if not os.path.exists(os.path.join(_REPO_DIR, f))]
    if miss:
        st.warning("명판/인감 파일이 자료 저장소(chjk-data)에 없습니다: " + ", ".join(miss)
                   + "  · PDF/명판이 비정상일 수 있습니다.")
    with st.spinner("불러오는 중…"):
        _fp = _data_fp()
        lt = _longoverdue_cached(_fp)
        book = _nyung_book(_fp)
    if not lt:
        st.info("현재 장기미수 거래처가 없습니다."); return
    qbiz = st.session_state.get("_routed_biz") or st.query_params.get("biz", "")
    labels = [f"[{o['reg']}] {o['name']} · {o['amt']:,}원" for o in lt]
    idx = next((i for i, o in enumerate(lt) if o["biz"] == qbiz), 0)
    # 팝업에서 새 거래처가 넘어오면 셀렉트박스를 그 거래처로 강제(이후엔 사용자 선택 유지)
    if qbiz and st.session_state.get("_ny_applied") != qbiz:
        st.session_state["_ny_applied"] = qbiz
        st.session_state["ny_sel"] = idx
    sel = st.selectbox("장기미수 거래처 선택", range(len(lt)), index=idx,
                       format_func=lambda i: labels[i], key="ny_sel")
    o = lt[sel]; biz = o["biz"]
    m = nyung.match_company(book, o["name"], biz) or {}
    st.caption("자동완성된 내용을 그대로 쓰거나 수정 후 생성하세요. 거래처를 바꾸면 자동으로 다시 채워집니다.")
    c1, c2 = st.columns([3, 2])
    corp = c1.text_input("거래처명 (정식 상호)", value=(m.get("name") or o["name"]), key=f"corp_{biz}")
    _amk = f"amt_{biz}"
    if _amk not in st.session_state:
        st.session_state[_amk] = f"{int(round(o['amt'])):,}"
    amt_raw = c2.text_input("미수액 (원)", key=_amk, on_change=_reformat_amt, args=(_amk,),
                            help="천 단위 쉼표가 자동으로 표시됩니다.")
    amt = int(re.sub(r"[^0-9]", "", amt_raw or "") or 0)
    c3, c4 = st.columns(2)
    rep = c3.text_input("대표자명 (수신 담당자)", value=m.get("rep", ""), key=f"rep_{biz}")
    tel = c4.text_input("수신 연락처", value=m.get("tel", ""), key=f"tel_{biz}")
    addr = st.text_input("수신 주소", value=m.get("addr", ""), key=f"addr_{biz}")
    reg = st.radio("관할 (명판·발신자·입금계좌가 바뀝니다)", ["서울", "화성"],
                   index=(0 if o["reg"] == "서울" else 1), horizontal=True, key=f"reg_{biz}")
    if not m:
        st.info("이 거래처는 명단에 없어 대표자·연락처·주소가 비어 있습니다. 직접 입력하세요.")
    st.caption(f"미수액 미리보기: {amt:,}원 · 하단 날짜는 다운로드 당일이 자동 입력됩니다.")
    if st.button("📄 내용증명 생성 (워드 + PDF)", type="primary"):
        if not corp.strip():
            st.error("거래처명을 입력하세요.")
        else:
            data = dict(거래처명=corp.strip(), 담당자=rep.strip(), 수신전화=tel.strip(),
                        수신주소=addr.strip(), 미수액=amt, 관할=reg)
            workdir = tempfile.mkdtemp(prefix="nyung_")
            safe = re.sub(r"[^가-힣A-Za-z0-9]", "_", corp.strip()) or "내용증명"
            docx_path = os.path.join(workdir, f"내용증명_{safe}.docx")
            try:
                with st.spinner("내용증명 생성 중… (PDF 하단 채움 최적화로 몇 초 더 걸릴 수 있어요)"):
                    nyung.make_docx(data, _REPO_DIR, docx_path)  # 워드: 맑은 고딕(이전 형식 그대로 유지)
                    wb = open(docx_path, "rb").read()
                    pdf_src = os.path.join(workdir, "pdf_src.docx")  # PDF: 별도 규칙(노토 + 하단까지 자동 채움)
                    pdf = _make_pdf_fit(data, pdf_src, workdir)
                    pb = open(pdf, "rb").read() if pdf else None
                st.session_state["ny_out"] = {"biz": biz, "word": wb, "pdf": pb, "name": f"내용증명_{safe}"}
            except Exception as e:
                st.session_state.pop("ny_out", None)
                st.error(f"문서 생성 오류: {e}")
    out = st.session_state.get("ny_out")
    if out and out.get("biz") == biz:
        st.success("생성 완료. 아래에서 내려받으세요.")
        d1, d2 = st.columns(2)
        d1.download_button("📥 워드(.docx) 받기", out["word"], file_name=out["name"] + ".docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_word")
        if out.get("pdf"):
            d2.download_button("📥 PDF 받기", out["pdf"], file_name=out["name"] + ".pdf",
                               mime="application/pdf", key="dl_pdf")
        else:
            d2.warning("PDF 변환 실패 — 서버에 libreoffice-writer 가 필요합니다. 워드는 정상입니다.")

# ===== 자료 처리 페이지 =====
def page_process():
    header(title="자료 처리")
    if ADMIN:
        st.markdown(f"<div style='margin:-6px 0 8px;'><span style='background:{NAVY};color:#fff;font-size:12px;font-weight:600;padding:3px 12px;border-radius:6px;'>🔑 관리자 계정</span></div>", unsafe_allow_html=True)
        st.caption("매출·매입 세금계산서, 어음수취내역, 은행거래내역을 올리면 자동으로 종류를 판별하고 기존 자료에 신규만 추가합니다.")
    if DATA_REPO and GIT_TOKEN and not _data_repo_dir():
        st.error("⚠ 자료 저장소(chjk-data) 연결 실패 — Secrets의 github_token 권한(chjk-data Contents: Read and write)과 github_data_repo 값을 확인하세요.")

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

    with st.expander("🛟 비상 복구 (직전본)", expanded=bool(st.session_state.get("confirm_restore"))):
        if os.path.exists(BACKUP_ZIP):
            st.caption("최근 갱신 직전 상태가 백업되어 있습니다. 문제가 생기면 한 번에 되돌릴 수 있어요.")
            bc1, bc2 = st.columns(2)
            bc1.download_button("직전본 다운로드 (서울·화성)", zip_backup_customers(), file_name="누적자료_직전본.zip", mime="application/zip", key="dl_bak")
            if ADMIN and bc2.button("⏪ 직전본으로 복구"):
                st.session_state["confirm_restore"] = True
            if ADMIN and st.session_state.get("confirm_restore"):
                st.warning("직전본으로 되돌리면 현재 누적본이 직전 상태로 바뀝니다. 정말 실행하시겠습니까?")
                rc1, rc2 = st.columns(2)
                if rc1.button("예, 복구 실행", type="primary", key="do_restore"):
                    st.session_state.pop("confirm_restore", None)
                    if store.restore_previous(DATA, DATA_ZIP, BACKUP_ZIP):
                        if GIT_TOKEN: store.git_commit_push(_REPO_DIR, GIT_TOKEN, DATA_REPO, "restore previous baseline")
                        st.session_state.pop("result", None)
                        st.success("직전본으로 복구했습니다."); st.rerun()
                if rc2.button("취소", key="cancel_restore"):
                    st.session_state.pop("confirm_restore", None); st.rerun()
        else:
            st.caption("아직 직전본 백업이 없습니다. (첫 갱신 후 생성됩니다)")

    if ADMIN:
        with st.expander("🔄 전체 다시 계산 (설정 변경 반영)"):
            st.caption("결제조건·별칭 등 설정만 바꿨을 때, 업로드 없이 전체 거래처의 상태·파일명을 현재 기준일로 다시 계산합니다.")
            if st.button("전체 다시 계산 실행", key="recalc_all"):
                old = {}
                for loc in ["서울", "화성"]:
                    for f in glob.glob(os.path.join(DATA, loc, "*.xlsx")):
                        nm = os.path.basename(f)
                        if nm.startswith("_"): continue
                        mb = re.search(r"(\d{3}-\d{2}-\d{5})", nm)
                        if mb:
                            old[mb.group(1)] = ("완납" if "완납" in nm else "장기미수" if "장기미수" in nm else "미수" if "미수" in nm else "진행")
                work = tempfile.mkdtemp(); out_dir = os.path.join(work, "out"); os.makedirs(out_dir, exist_ok=True)
                with st.spinner("전체 거래처 다시 계산 중… (수십 초)"):
                    try:
                        res = pipeline.process({"서울": [], "화성": []}, DATA, out_dir, ref_date=basis_date())
                        ok, probs = pipeline.verify(out_dir, DATA)
                    except Exception as e:
                        ok = False; probs = [str(e)]
                if ok:
                    changes = [(s[0], s[1], old.get(s[2], "?"), s[3]) for s in res["summary"] if old.get(s[2], s[3]) != s[3]]
                    store.apply_update(out_dir, DATA, DATA_ZIP, BACKUP_ZIP)
                    saved, msg = store.git_commit_push(_REPO_DIR, GIT_TOKEN, DATA_REPO, "recompute all statuses")
                    st.cache_data.clear()
                    note = "  (GitHub 영구저장 완료)" if saved else f"  (GitHub 미저장: {msg})"
                    if changes:
                        st.success(f"전체 다시 계산 완료 — {len(changes)}곳 상태가 바뀌었습니다." + note)
                        cdf = pd.DataFrame(changes, columns=["지역", "거래처", "이전", "변경"]).sort_values(["지역", "거래처"]).reset_index(drop=True)
                        st.dataframe(cdf, use_container_width=True, hide_index=True)
                    else:
                        st.success("전체 다시 계산 완료 — 바뀐 상태가 없습니다." + note)
                else:
                    st.error("재계산 검증 실패 — 반영하지 않았습니다."); st.write(probs[:10])

        _md = st.session_state.pop("manual_done", None)
        with st.expander("✏️ 거래처 파일 직접 수정 후 업로드 (수동 교체)", expanded=bool(_md)):
            if _md: st.success(_md)
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
                    saved, msg = store.git_commit_push(_REPO_DIR, GIT_TOKEN, DATA_REPO, "manual edit replace")
                    st.session_state.pop("result", None)
                    st.session_state["manual_done"] = f"✅ {len(good)}개 거래처 교체 완료." + ("  (GitHub 영구저장 완료)" if saved else f"  (GitHub 미저장: {msg})")
                    st.toast(f"{len(good)}개 거래처 교체 완료 ✅")
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
                    res = pipeline.process(uploads, DATA if os.path.isdir(DATA) else up_dir, out_dir, progress=prog, ref_date=basis_date())
                except Exception as e:
                    st.error(f"처리 오류: {e}"); st.stop()
                ok, probs = pipeline.verify(out_dir, DATA)
                if ok:
                    store.apply_update(out_dir, DATA, DATA_ZIP, BACKUP_ZIP)
                    open(BASIS, "w", encoding="utf-8").write(res.get("basis") or datetime.date.today().isoformat())
                    saved, msg = store.git_commit_push(_REPO_DIR, GIT_TOKEN, DATA_REPO, "update baseline via app")
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

if not check_pw(): st.stop()
ADMIN = st.session_state.get("role") == "admin"
# 페이지 객체(전역) — 팝업 트리거와 '돌아가기'에서 st.switch_page 로 이동. 미수현황 탭은 항상 홈으로 복귀.
_dashboard_page = st.Page(page_dashboard, title="미수 현황", icon="📊", default=True)
_sales_page = st.Page(page_sales, title="매출 현황", icon="📈")
_process_page = st.Page(page_process, title="자료 처리", icon="🗂")
_nyung_page = st.Page(page_nyung, title="내용증명", icon="📄", url_path="page_nyung")
_pages = [_dashboard_page, _sales_page, _nyung_page]   # 내용증명은 관리자·실무자 모두(탭은 CSS로 숨김, 팝업으로 진입)
_pages.append(_process_page)
_pg = st.navigation(_pages, position="top")
_pg.run()
