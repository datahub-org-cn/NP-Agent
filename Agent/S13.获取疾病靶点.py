# -*- coding: utf-8 -*-
"""
S13.获取疾病靶点.py
===================
从 Open Targets Platform (公开 GraphQL API) 获取 S13疾病.xlsx 中列出的
疾病的关联靶点(含 Gene symbol)。

接口: https://api.platform.opentargets.org/api/v4/graphql
     disease(efoId:"MONDO_xxxxx").associatedTargets(...)  按关联评分降序分页
纳入阈值: 关联评分 ≥ OT_SCORE_MIN (默认 0.1), 低于阈值的关联靶点不纳入

疾病→ID 解析(仅 Open Targets):
    1. 若 S13疾病.xlsx 第2列为 Open Targets ID (MONDO_/EFO_开头), 直接使用;
    2. 否则在内置已知疾病映射表中匹配(归一化后不区分大小写);
    3. 均未命中 -> 提示在表中补充 ID, 跳过该疾病。

输入: S13疾病.xlsx (列1: 疾病; 可选列2: Open Targets ID)
输出: S13.疾病靶点.xlsx
      - 数据说明 : 疾病 / Open Targets ID / 关联靶点总数 / 含Gene symbol靶点数 / 获取日期
      - 靶点详情 : 疾病 / Gene symbol / Ensembl ID / 蛋白名 / 基因类型 / 关联评分 / novelty / 主要证据来源

用法: python S13.获取疾病靶点.py
"""
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request

API = "https://api.platform.opentargets.org/api/v4/graphql"
INPUT = "S13疾病.xlsx"
XLSX_OUT = "S13.疾病靶点.xlsx"
PAGE_SIZE = 100
OT_SCORE_MIN = 0.1   # Open Targets 关联评分阈值(≥0.1 才纳入)

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

# 内置已知疾病映射(归一化名 -> Open Targets 疾病ID)
KNOWN_DISEASES = {
    "depression": "MONDO_0002050",
    "depressivedisorder": "MONDO_0002050",
    "抑郁障碍": "MONDO_0002050",
    "抑郁症": "MONDO_0002050",
}


def norm(s):
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", str(s).lower())


def read_diseases(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        name = str(row[0]).strip()
        if not name:
            continue
        did = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        out.append({"疾病": name, "ID": did})
    return out


def gql(query):
    body = json.dumps({"query": query}).encode("utf-8")
    req = urllib.request.Request(API, data=body,
                                 headers={"Content-Type": "application/json", **UA})
    for attempt in range(4):
        try:
            with urllib.request.urlopen(req, timeout=120) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 429:
                time.sleep(5 * (attempt + 1))
                continue
            time.sleep(2 * (attempt + 1))
        except Exception:
            time.sleep(2 * (attempt + 1))
    raise RuntimeError("Open Targets GraphQL 请求失败")


def fetch_page(efo_id, index):
    q = """{
      disease(efoId: "%s") {
        id name
        associatedTargets(page: {index: %d, size: %d}, orderByScore: "desc") {
          count
          rows { score novelty datasourceScores { id score } target { id approvedSymbol approvedName biotype } }
        }
      }
    }""" % (efo_id, index, PAGE_SIZE)
    res = gql(q)
    dis = (res.get("data") or {}).get("disease")
    if dis is None:
        errs = res.get("errors") or []
        raise RuntimeError(f"疾病 {efo_id} 查询失败: {errs}")
    return dis


def main():
    if not os.path.exists(INPUT):
        print(f"错误: 找不到 {INPUT}")
        sys.exit(1)
    diseases = read_diseases(INPUT)
    print(f"读取到 {len(diseases)} 个疾病: {[d['疾病'] for d in diseases]}")

    summary_rows, target_rows = [], []
    for di, d in enumerate(diseases, 1):
        name = d["疾病"]
        did = d["ID"].strip()
        if not did:
            did = KNOWN_DISEASES.get(norm(name), "")
        if not did:
            print(f"[{di}/{len(diseases)}] {name}: 未找到 Open Targets ID, 跳过(请在S13疾病.xlsx第2列补充MONDO/EFO ID)")
            summary_rows.append({"疾病": name, "Open Targets ID": "", "关联靶点总数": "",
                                 "含Gene symbol靶点数": 0, "获取日期": time.strftime("%Y-%m-%d")})
            continue
        print(f"[{di}/{len(diseases)}] 获取 {name} ({did}) 靶点 ...", flush=True)
        first = fetch_page(did, 0)
        total = first["associatedTargets"]["count"]
        print(f"   关联靶点总数: {total}", flush=True)

        rows, seen = [], set()
        n_pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        for idx in range(n_pages):
            dd = fetch_page(did, idx)
            for r in dd["associatedTargets"]["rows"]:
                t = r["target"]
                ens = (t.get("id") or "").strip()
                if ens in seen:
                    continue
                seen.add(ens)
                symbol = (t.get("approvedSymbol") or "").strip()
                if not symbol:
                    continue
                try:
                    scf = float(r.get("score") or 0)
                except (TypeError, ValueError):
                    scf = 0.0
                if scf < OT_SCORE_MIN:
                    continue
                ds = sorted(r.get("datasourceScores", []), key=lambda x: -(x.get("score") or 0))
                sources = "; ".join(x["id"] for x in ds[:3])
                rows.append({
                    "疾病": name, "Gene symbol": symbol,
                    "Ensembl ID": t.get("id", ""), "蛋白名": t.get("approvedName", ""),
                    "基因类型": t.get("biotype", ""), "关联评分": r.get("score", ""),
                    "novelty": r.get("novelty", ""), "主要证据来源": sources,
                })
            if (idx + 1) % 5 == 0 or idx == n_pages - 1:
                print(f"   分页 {idx + 1}/{n_pages}, 已收集 {len(rows)}", flush=True)
            time.sleep(0.3)

        rows.sort(key=lambda r: float(r["关联评分"] or 0), reverse=True)
        target_rows.extend(rows)
        summary_rows.append({"疾病": name, "Open Targets ID": did, "关联靶点总数": total,
                             "含Gene symbol靶点数": len(rows), "获取日期": time.strftime("%Y-%m-%d")})
        print(f"   含 Gene symbol 靶点: {len(rows)}", flush=True)

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="4472C4")
    head_font = Font(bold=True, color="FFFFFF")

    def style_header(ws, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=1, column=c)
            cell.fill = head_fill
            cell.font = head_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    ws0 = wb.active
    ws0.title = "数据说明"
    ws0.append(["疾病", "Open Targets ID", "关联靶点总数", "含Gene symbol靶点数", "获取日期"])
    for r in summary_rows:
        ws0.append([r["疾病"], r["Open Targets ID"], r["关联靶点总数"],
                    r["含Gene symbol靶点数"], r["获取日期"]])
    style_header(ws0, 5)
    for i, c in enumerate(["疾病", "Open Targets ID", "关联靶点总数", "含Gene symbol靶点数", "获取日期"], 1):
        ws0.column_dimensions[get_column_letter(i)].width = 18

    ws1 = wb.create_sheet("靶点详情")
    cols = ["疾病", "Gene symbol", "Ensembl ID", "蛋白名", "基因类型",
            "关联评分", "novelty", "主要证据来源"]
    ws1.append(cols)
    for r in target_rows:
        ws1.append([r[c] for c in cols])
    style_header(ws1, len(cols))
    for i, w in enumerate([16, 16, 18, 50, 14, 10, 10, 30], 1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.freeze_panes = "A2"

    wb.save(XLSX_OUT)
    print(f"已写出 {XLSX_OUT}: 靶点详情 {len(target_rows)} 行")


if __name__ == "__main__":
    main()
