# -*- coding: utf-8 -*-
"""매출 현황: 매입·매출 계산서 기준 매출/매입·판매순위 HTML 렌더.
 - 과거자료 _매출자료.json (2022~2026.4)  → 기준선(boundary) 이하 월에 사용
 - 누적분 _매출집계.json (자료 처리 때 세금계산서 누적, 승인번호 dedup) → boundary 이후 월(2026.5~) 자동 반영
 두 자료를 합쳐 years/summary/rank 구조로 만들어 sales_template.html 에 주입한다."""
import os, json, re


def _cname(s):
    s = str(s).replace("（", "(").replace("）", ")")
    s = re.sub(r"\(주\)|주식회사|㈜|\(유\)|유한회사|\(자\)|영업소|前", "", s)
    s = re.sub(r"\([^)]*\)", "", s)
    return re.sub(r"[^가-힣A-Za-z0-9]", "", s).lower()


def load_data(data_dir):
    """과거 기준자료(_매출자료.json) 로드. 없으면 None."""
    p = os.path.join(data_dir, "_매출자료.json")
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as fh:
        return json.load(fh)


def _boundary(base):
    """과거자료의 마지막 채워진 월 → 'YYYY-MM' (예: '2026-04'). 이 월까지는 과거자료가 기준."""
    best = ""
    for y, months in (base.get("years") or {}).items():
        for e in months:
            if e.get("매출합"):
                try: mi = int(str(e["월"]).replace("월", ""))
                except Exception: continue
                ym = f"{y}-{mi:02d}"
                if ym > best: best = ym
    return best or "0000-00"


def _agg_store(store, boundary):
    """누적분에서 boundary 이후 월만 집계. 반환: (월별집계, 거래처순위집계)."""
    month = {}   # (year, m_int, reg) -> {'매출':n,'매입':n}
    rank = {}    # (year, reg) -> {cname: [display, sup]}
    for typ in ("매출", "매입"):
        for rec in (store.get(typ) or {}).values():
            ym = rec.get("ym", "")
            if len(ym) < 7 or ym <= boundary:
                continue
            reg = rec.get("reg")
            if reg not in ("서울", "화성"):
                continue
            y = ym[:4]; mi = int(ym[5:7]); sup = rec.get("sup", 0) or 0
            d = month.setdefault((y, mi, reg), {"매출": 0, "매입": 0})
            d[typ] += sup
            if typ == "매출":
                nm = rec.get("name") or rec.get("biz") or ""
                cn = _cname(nm) or nm
                rk = rank.setdefault((y, reg), {})
                if cn in rk: rk[cn][1] += sup
                else: rk[cn] = [nm, sup]
    return month, rank


def _merge_rank(entries_lists):
    """[{업체,매출액}...] 들을 정규화명(cname) 기준으로 합산 → 정렬된 리스트."""
    m = {}
    for lst in entries_lists:
        for z in lst:
            cn = _cname(z["업체"]) or z["업체"]
            if cn in m: m[cn][1] += z["매출액"]
            else: m[cn] = [z["업체"], z["매출액"]]
    return sorted(({"업체": v[0], "매출액": v[1]} for v in m.values()), key=lambda r: -r["매출액"])


def _merge_named(lists, field):
    """[{업체, <field>}...] 들을 정규화명(cname) 기준으로 합산 → 정렬된 리스트."""
    m = {}
    for lst in lists:
        for z in lst:
            cn = _cname(z["업체"]) or z["업체"]
            if cn in m: m[cn][1] += z[field]
            else: m[cn] = [z["업체"], z[field]]
    return sorted(({"업체": v[0], field: v[1]} for v in m.values()), key=lambda r: -r[field])


def _clean_corp(s):
    """표시용 이름에서 법인 표기((주)·주식회사·㈜·(유)·(자) 등) 제거. (오류동) 같은 지점 표기는 유지."""
    s = str(s or "").replace("（", "(").replace("）", ")")
    s = re.sub(r"주식회사|유한회사|합자회사|㈜|\(주\)|\(유\)|\(자\)", "", s)
    return re.sub(r"\s+", " ", s).strip()


def _disp_buy(name, rep, biz):
    """매입처 표시명: 상호(법인표기 제거). 상호 없으면 '대표자 (사업자번호)', 그것도 없으면 사업자번호."""
    nm = _clean_corp(name)
    if nm and nm.lower() != "nan":
        return nm
    rep = str(rep or "").strip()
    return f"{rep} ({biz})" if rep and rep.lower() != "nan" else str(biz)


def build_data(data_dir):
    """과거자료 + 누적분을 합친 매출 현황 데이터(years/summary/rank)."""
    base = load_data(data_dir)
    if base is None:
        return None
    store = {}
    sp = os.path.join(data_dir, "_매출집계.json")
    if os.path.exists(sp):
        try:
            with open(sp, encoding="utf-8") as fh: store = json.load(fh)
        except Exception:
            store = {}
    boundary = _boundary(base)
    month_agg, rank_agg = _agg_store(store, boundary)

    # 1) 월별: 과거자료 복사 + 누적분(boundary 이후) 추가
    years = {y: [dict(m) for m in months] for y, months in (base.get("years") or {}).items()}
    for (y, mi, reg), d in month_agg.items():
        lst = years.setdefault(y, [])
        label = f"{mi}월"
        ent = next((e for e in lst if e.get("월") == label), None)
        if ent is None:
            ent = {"월": label, "매출서울": 0, "매출화성": 0, "매입서울": 0, "매입화성": 0, "매출합": 0, "매입합": 0, "총이익": 0}
            lst.append(ent)
        ent["매출" + reg] = d["매출"]
        ent["매입" + reg] = d["매입"]
    for y, lst in years.items():
        for e in lst:
            e["매출합"] = (e.get("매출서울") or 0) + (e.get("매출화성") or 0)
            e["매입합"] = (e.get("매입서울") or 0) + (e.get("매입화성") or 0)
            e["총이익"] = e["매출합"] - e["매입합"]
        lst.sort(key=lambda e: int(str(e["월"]).replace("월", "")))

    # 2) 연도 요약(월별에서 재계산)
    summary = {"매출 (계산서 기준)": {}, "매입 (계산서 기준)": {}, "매출총이익": {}, "마진율": {}, "전년 대비 매출 증감": {}}
    prev = None
    for y in sorted(years.keys()):
        sa = sum(e["매출합"] for e in years[y]); bu = sum(e["매입합"] for e in years[y]); pr = sa - bu
        summary["매출 (계산서 기준)"][y] = sa
        summary["매입 (계산서 기준)"][y] = bu
        summary["매출총이익"][y] = pr
        summary["마진율"][y] = (pr / sa) if sa else 0
        summary["전년 대비 매출 증감"][y] = ((sa - prev) / prev) if prev else None
        prev = sa

    # 3) 거래처 순위: 과거 서울/화성 + 누적분 병합 → 전체·전기간 파생
    rank = {}
    yrs = sorted(years.keys())
    for y in yrs:
        for reg in ("서울", "화성"):
            base_lst = (base.get("rank") or {}).get(f"{y}|{reg}", [])
            acc = [{"업체": v[0], "매출액": v[1]} for v in rank_agg.get((y, reg), {}).values()]
            rank[f"{y}|{reg}"] = _merge_rank([base_lst, acc])
        rank[f"{y}|전체"] = _merge_rank([rank[f"{y}|서울"], rank[f"{y}|화성"]])
    for reg in ("서울", "화성", "전체"):
        rank[f"전체|{reg}"] = _merge_rank([rank[f"{y}|{reg}"] for y in yrs])

    # 4) 매입처 순위: 과거 _매입자료.json + 누적 매입(업체명 있는 분, boundary 이후) 병합
    hbuy = {}
    bp = os.path.join(data_dir, "_매입자료.json")
    if os.path.exists(bp):
        try:
            with open(bp, encoding="utf-8") as fh: hbuy = (json.load(fh) or {}).get("buyrank", {})
        except Exception:
            hbuy = {}
    buyacc = {}
    for rec in (store.get("매입") or {}).values():
        ym = rec.get("ym", "")
        if len(ym) < 7 or ym <= boundary:
            continue
        reg = rec.get("reg")
        if reg not in ("서울", "화성"):
            continue
        biz = str(rec.get("biz") or "")
        nm = _disp_buy(rec.get("name"), rec.get("rep"), biz)
        key = biz or nm
        d = buyacc.setdefault((ym[:4], reg), {})
        if key in d: d[key][1] += rec.get("sup", 0)
        else: d[key] = [nm, rec.get("sup", 0)]
    buyrank = {}
    for y in yrs:
        for reg in ("서울", "화성"):
            base_lst = [{"업체": _clean_corp(z["업체"]) or z["업체"], "매입액": z["매입액"]} for z in hbuy.get(f"{y}|{reg}", [])]
            acc = [{"업체": v[0], "매입액": v[1]} for v in buyacc.get((y, reg), {}).values()]
            buyrank[f"{y}|{reg}"] = _merge_named([base_lst, acc], "매입액")
        buyrank[f"{y}|전체"] = _merge_named([buyrank[f"{y}|서울"], buyrank[f"{y}|화성"]], "매입액")
    for reg in ("서울", "화성", "전체"):
        buyrank[f"전체|{reg}"] = _merge_named([buyrank[f"{y}|{reg}"] for y in yrs], "매입액")

    return {"years": years, "summary": summary, "rank": rank, "buyrank": buyrank}


def _basis_label(data):
    """최신 연도·마지막으로 채워진 월 → '2026년 4월'."""
    yrs = sorted((data.get("years") or {}).keys())
    for y in reversed(yrs):
        last = None
        for m in data["years"][y]:
            if m.get("매출합"):
                last = m.get("월")
        if last:
            return f"{y}년 {last}"
    return ""


def render(data_dir, admin=False, template_path=None, basis_iso=None):
    """매출 현황 HTML 문자열 반환. data_dir 에 _매출자료.json 필요(_매출집계.json 있으면 자동 반영).
    basis_iso: 다른 탭과 동일하게 표시할 기준일(YYYY-MM-DD). 없으면 매출자료의 마지막 월."""
    data = build_data(data_dir)
    if data is None:
        return None
    tp = template_path or os.path.join(os.path.dirname(os.path.abspath(__file__)), "sales_template.html")
    tpl = open(tp, encoding="utf-8").read()
    label = ""
    if basis_iso:
        import datetime
        try:
            d = datetime.date.fromisoformat(basis_iso)
            label = f"{d.year}년 {d.month}월 {d.day}일"
        except Exception:
            label = ""
    if not label:
        label = _basis_label(data)
    badge = ("<div style='margin:-10px 0 8px'><span style='background:#1B3A6B;color:#fff;font-size:12px;"
             "font-weight:600;padding:3px 12px;border-radius:6px'>🔑 관리자 계정</span></div>") if admin else ""
    return (tpl.replace("@@DATA@@", json.dumps(data, ensure_ascii=False))
               .replace("@@SBASIS@@", label)
               .replace("@@ADMINBADGE@@", badge))
