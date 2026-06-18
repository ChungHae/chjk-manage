# -*- coding: utf-8 -*-
"""매출 현황: 매입·매출 계산서 기준 매출/매입·판매순위.
과거자료(_매출자료.json, 2022~2026.4)를 읽어 HTML로 렌더한다.
2026.5+ 자동계산(거래처 파일/세금계산서 기반)은 build_data에서 확장 예정."""
import os, json


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


def load_data(data_dir):
    p = os.path.join(data_dir, "_매출자료.json")
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as fh:
        return json.load(fh)


def render(data_dir, admin=False, template_path=None):
    """매출 현황 HTML 문자열 반환. data_dir 에 _매출자료.json 필요."""
    data = load_data(data_dir)
    if data is None:
        return None
    tp = template_path or os.path.join(os.path.dirname(os.path.abspath(__file__)), "sales_template.html")
    tpl = open(tp, encoding="utf-8").read()
    badge = ("<div class=abadge>🔑 관리자 계정</div>") if admin else ""
    return (tpl.replace("@@DATA@@", json.dumps(data, ensure_ascii=False))
               .replace("@@SBASIS@@", _basis_label(data))
               .replace("@@ADMINBADGE@@", badge))
