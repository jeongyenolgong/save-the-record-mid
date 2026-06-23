#!/usr/bin/env python3
# 이 레포의 콘텐츠 엑셀(content/*.xlsx) → content.json 변환
# 사용법:  python3 convert.py
import json, os, glob
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SECTION_ORDER = ["생산 배경","생산 방법","생존 이야기","기록문화적 가치","현재 보존 방법","접근·열람 방법"]
METHOD = {"web":"웹제공", "card":"카드"}

def cols(ws):
    return {ws.cell(1,c).value: c for c in range(1, ws.max_column+1)}

def find_xlsx():
    xs = [x for x in glob.glob(os.path.join(HERE,"content","*.xlsx"))
          if not os.path.basename(x).startswith("~$")]
    if not xs:
        raise SystemExit("content/ 폴더에 .xlsx 가 없습니다.")
    return sorted(xs)[0]

def convert(xlsx_path, json_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    it, co = wb["items"], wb["content"]
    ic, cc = cols(it), cols(co)
    meta = {}
    for r in range(2, it.max_row+1):
        no = it.cell(r, ic["no"]).value
        if no in (None,""): continue
        no = int(no)
        gi = lambda k: it.cell(r, ic[k]).value
        meta[no] = dict(
            no=no, category=gi("category") or "",
            created=int(gi("created")) if gi("created") not in (None,"") else None,
            era_label=str(gi("era_label") or ""), name=gi("name") or "",
            country="한국",
            year=int(gi("year")) if gi("year") not in (None,"") else None,
            difficulty=int(gi("difficulty")) if gi("difficulty") not in (None,"") else None,
            desc=gi("desc") or "", items={})
    for r in range(2, co.max_row+1):
        no = co.cell(r, cc["no"]).value
        if no in (None,"") or int(no) not in meta: continue
        no = int(no)
        label = co.cell(r, cc["section_label"]).value
        code = co.cell(r, cc["code"]).value
        code = "" if code in (None,"") else str(int(code) if isinstance(code,(int,float)) else code)
        meta[no]["items"][label] = dict(
            method=METHOD.get(co.cell(r, cc["method"]).value, "웹제공"), code=code,
            content=co.cell(r, cc["content"]).value or "")
    out = []
    for no in sorted(meta):
        m = meta[no]
        m["items"] = {k: m["items"][k] for k in SECTION_ORDER if k in m["items"]}
        out.append(m)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    filled = sum(1 for m in out for s in m["items"].values() if s["content"])
    print(f"{os.path.basename(json_path)} 생성: 유산 {len(out)}건, 본문 {filled}개")

if __name__ == "__main__":
    convert(find_xlsx(), os.path.join(HERE, "content.json"))
