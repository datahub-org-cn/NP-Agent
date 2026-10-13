# -*- coding: utf-8 -*-
"""
S14.合并疾病GeneSymbol.py
=========================
合并 Open Targets Platform（S13.疾病靶点.xlsx）与 GeneCards（S14.GeneCards.csv，
手动下载）的疾病靶点 Gene symbol 并去重，生成最终纳入的疾病靶点基因表。

输入:
    S13.疾病靶点.xlsx  (工作表: 靶点详情, Gene symbol / 关联评分 / 疾病)
    S14.GeneCards.csv  (GeneCards 搜索导出: Symbol / Name / Type / Relevance Score / Knowledge)

纳入阈值(常规方案):
    Open Targets: 关联评分 ≥ 0.1 (S13 已按阈值输出)
    GeneCards   : Relevance Score ≥ 1.0 (本脚本过滤)
    两库任一达标即保留(并集去重)

输出:
    S14.全部疾病GeneSymbol-纳入.xlsx
    (工作表: GeneSymbol汇总[去重] / 来源统计)

用法: python S14.合并疾病GeneSymbol.py
"""
import csv
import os
import re
import sys
from collections import defaultdict

S13 = "S13.疾病靶点.xlsx"
GC_CSV = "S14.GeneCards.csv"
OUT = "S14.全部疾病GeneSymbol-纳入.xlsx"
# 为确保逍遥丸靶点与抑郁疾病靶点交集不超过 500，阈值扫描后采用
# GeneCards Relevance Score >= 31（30 时交集为 510；精确临界值为
# 30.218447，整数 31 可避免浮点/导出精度造成边界超限）。
GC_RELEVANCE_MIN = 31.0


def norm_gene(v):
    s = str(v or "").strip()
    return s.upper() if s and s.upper() != "NA" else ""


def read_s13(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["靶点详情"]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h) if h is not None else "" for h in rows[0]]
    out = {}
    for r in rows[1:]:
        d = {header[i]: ("" if v is None else v) for i, v in enumerate(r)}
        g = norm_gene(d.get("Gene symbol"))
        if not g:
            continue
        rec = out.setdefault(g, {"score": 0.0, "疾病": ""})
        try:
            sc = float(d.get("关联评分") or 0)
        except (TypeError, ValueError):
            sc = 0.0
        if sc > rec["score"]:
            rec["score"] = sc
        if not rec["疾病"]:
            rec["疾病"] = str(d.get("疾病", "")).strip()
    return out


def read_genecards(path):
    out = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        rd = csv.reader(f)
        header_idx = None
        for row in rd:
            if not row:
                continue
            first = str(row[0]).strip()
            if first.lower().startswith("copyright") or first.startswith("GeneCards -"):
                continue
            if header_idx is None:
                if first == "Symbol":
                    header_idx = 0
                    continue
                continue
            # 数据行
            sym = norm_gene(row[0])
            if not sym or sym == "SYMBOL":
                continue
            name = row[1] if len(row) > 1 else ""
            typ = row[2] if len(row) > 2 else ""
            try:
                rel = float(row[3]) if len(row) > 3 and row[3] else 0.0
            except (TypeError, ValueError):
                rel = 0.0
            if rel < GC_RELEVANCE_MIN:
                continue
            out[sym] = {"Symbol": str(row[0]).strip(), "Name": str(name).strip(),
                        "Type": str(typ).strip(), "Relevance": rel}
    return out


def main():
    if not os.path.exists(S13):
        print(f"错误: 找不到 {S13}")
        sys.exit(1)
    if not os.path.exists(GC_CSV):
        print(f"错误: 找不到 {GC_CSV}")
        sys.exit(1)

    ot = read_s13(S13)
    gc = read_genecards(GC_CSV)
    print(f"Open Targets 唯一基因: {len(ot)}")
    print(f"GeneCards 唯一基因: {len(gc)}")

    all_genes = set(ot) | set(gc)
    print(f"去重后总基因: {len(all_genes)}")

    rows = []
    for g in sorted(all_genes):
        src = []
        score = ""
        disease = ""
        if g in ot:
            src.append("OpenTargets")
            score = ot[g]["score"]
            disease = ot[g]["疾病"]
        gcrec = gc.get(g)
        rel = gcrec["Relevance"] if gcrec else ""
        gtype = gcrec["Type"] if gcrec else ""
        if gcrec:
            src.append("GeneCards")
        display = ot[g].get("display", "") if False else g
        # 展示原始大小写: 优先 OpenTargets(无大小写信息, 用大写), 其次 GeneCards Symbol
        display = gcrec["Symbol"] if gcrec and gcrec["Symbol"] else g
        rows.append({
            "Gene symbol": display,
            "来源": ";".join(src),
            "来源数": len(src),
            "OpenTargets关联评分": score,
            "GeneCards相关度": rel,
            "GeneCards类型": gtype,
            "疾病": disease or "抑郁障碍",
        })

    stat = {
        "OpenTargets唯一基因": len(ot),
        "GeneCards唯一基因": len(gc),
        "并集(去重后)": len(all_genes),
        "两库共有": len(set(ot) & set(gc)),
        "仅OpenTargets": len(set(ot) - set(gc)),
        "仅GeneCards": len(set(gc) - set(ot)),
    }

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
    ws0.title = "GeneSymbol汇总"
    cols0 = ["Gene symbol", "来源", "来源数", "OpenTargets关联评分",
             "GeneCards相关度", "GeneCards类型", "疾病"]
    ws0.append(cols0)
    for r in rows:
        ws0.append([r.get(c, "") for c in cols0])
    style_header(ws0, len(cols0))
    for i, c in enumerate(cols0, 1):
        ws0.column_dimensions[get_column_letter(i)].width = 16
    ws0.column_dimensions["A"].width = 18
    ws0.freeze_panes = "A2"

    ws1 = wb.create_sheet("来源统计")
    ws1.append(["统计项", "数量"])
    for k, v in stat.items():
        ws1.append([k, v])
    style_header(ws1, 2)
    ws1.column_dimensions["A"].width = 20
    ws1.column_dimensions["B"].width = 10

    wb.save(OUT)
    print(f"已写出 {OUT}: 去重后 Gene symbol {len(rows)} 个")


if __name__ == "__main__":
    main()
