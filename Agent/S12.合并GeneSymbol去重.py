# -*- coding: utf-8 -*-
"""
S12.合并GeneSymbol去重.py
========================
汇总数据库靶点(S6/S7/S8/S9)与预测靶点(S11)的全部 Gene symbol 并去重,
供后续 Venn(核心靶点)与调控网络共同使用:

    S6 : S6.药物级靶点_HERB.xlsx      (药物级-HERB, 含Gene symbol)
    S7 : S7.成分级靶点_HERB.xlsx      (成分级-HERB, 含Gene symbol)
    S8 : S8.药物级靶点_TCMSP-补全GeneSymbol.xlsx (药物级-TCMSP)
    S9 : S9.成分级靶点_TCMSP-补全GeneSymbol.xlsx (成分级-TCMSP)
    S11: S11.预测靶点_SwissTargetPrediction.xlsx (STP预测靶点)

输出: S12.全部GeneSymbol汇总.xlsx
    (工作表: GeneSymbol汇总[去重后] / 来源统计)

用法: python S12.合并GeneSymbol去重.py
"""
import os
import re
import sys
from collections import defaultdict

S6 = "S6.药物级靶点_HERB.xlsx"
S7 = "S7.成分级靶点_HERB.xlsx"
S8 = "S8.药物级靶点_TCMSP-补全GeneSymbol.xlsx"
S9 = "S9.成分级靶点_TCMSP-补全GeneSymbol.xlsx"
S11 = "S11.预测靶点_SwissTargetPrediction.xlsx"
OUT = "S12.全部GeneSymbol汇总.xlsx"

DRUGS = ["黄芪", "柴胡", "川芎", "水蛭", "枳实", "川牛膝", "莪术",
         "三棱", "土鳖虫", "大黄", "土茯苓", "槟榔", "赶黄草"]


def read_xlsx(path, sheet):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h) if h is not None else "" for h in rows[0]]
    out = []
    for r in rows[1:]:
        if r is None or all(v is None or str(v).strip() == "" for v in r):
            continue
        out.append({header[i]: ("" if v is None else v) for i, v in enumerate(r)})
    wb.close()
    return out


def norm_gene(v):
    s = str(v or "").strip()
    return s.upper() if s and s.upper() != "NA" else ""


def collect(path, sheet, drug_col=None):
    """返回 {gene: {'drugs':set, 'uniprot':str, 'n':int}}"""
    rows = read_xlsx(path, sheet)
    out = defaultdict(lambda: {"drugs": set(), "uniprot": "", "n": 0})
    for r in rows:
        g = norm_gene(r.get("Gene symbol"))
        if not g:
            continue
        rec = out[g]
        if drug_col:
            d = str(r.get(drug_col, "")).strip()
            if d:
                rec["drugs"].add(d)
        elif sheet in DRUGS:
            rec["drugs"].add(sheet)
        u = str(r.get("Uniprot ID", "") or r.get("UniProt ID", "") or "").strip()
        if u and u.upper() != "NA" and not rec["uniprot"]:
            rec["uniprot"] = u
        rec["n"] += 1
    return out


def sort_drugs(ds):
    return sorted(ds, key=lambda d: DRUGS.index(d) if d in DRUGS else 999)


def main():
    if not all(os.path.exists(p) for p in (S8, S9, S11)):
        print("错误: 缺少输入文件")
        for p in (S8, S9, S11):
            if not os.path.exists(p):
                print("  找不到:", p)
        sys.exit(1)

    # 数据库靶点: HERB(S6/S7) + TCMSP(S8/S9)
    genes_s6 = collect(S6, "靶点明细", drug_col="药物")
    genes_s7 = collect(S7, "靶点明细", drug_col="药物")
    genes_s8 = collect(S8, "靶点明细", drug_col="药物")
    genes_s9 = collect(S9, "靶点明细", drug_col="药物")
    # S11: 逐药物工作表(水蛭/土鳖虫/赶黄草), 药物=工作表名
    genes_s11 = defaultdict(lambda: {"drugs": set(), "uniprot": "", "n": 0})
    from openpyxl import load_workbook
    wb11 = load_workbook(S11, read_only=True, data_only=True)
    for sheet in wb11.sheetnames:
        if sheet in ("药物汇总", "全部靶点汇总"):
            continue
        part = collect(S11, sheet)
        for g, rec in part.items():
            genes_s11[g]["drugs"].update(rec["drugs"] or {sheet})
            genes_s11[g]["uniprot"] = genes_s11[g]["uniprot"] or rec["uniprot"]
            genes_s11[g]["n"] += rec["n"]
    wb11.close()

    print(f"S6 唯一基因: {len(genes_s6)}")
    print(f"S7 唯一基因: {len(genes_s7)}")
    print(f"S8 唯一基因: {len(genes_s8)}")
    print(f"S9 唯一基因: {len(genes_s9)}")
    print(f"S11 唯一基因: {len(genes_s11)}")

    all_genes = (set(genes_s6) | set(genes_s7) | set(genes_s8)
                 | set(genes_s9) | set(genes_s11))
    print(f"去重后总基因: {len(all_genes)}")

    rows = []
    for g in sorted(all_genes):
        src = []
        drugs = set()
        uniprot = ""
        for tag, gm in (("S6", genes_s6), ("S7", genes_s7), ("S8", genes_s8),
                        ("S9", genes_s9), ("S11", genes_s11)):
            if g in gm:
                src.append(tag)
                drugs |= gm[g]["drugs"]
                uniprot = uniprot or gm[g]["uniprot"]
        # 展示用原始大小写: 优先取 S6/S7/S8/S9 中的写法
        display = g
        for source_map in (genes_s6, genes_s7, genes_s8, genes_s9):
            for k in source_map:
                if k.upper() == g:
                    display = k
                    break
            else:
                continue
            break
        rows.append({
            "Gene symbol": display,
            "来源": ";".join(src),
            "来源数": len(src),
            "涉及药物": "、".join(sort_drugs(drugs)),
            "药物数": len(drugs),
            "UniProt ID": uniprot,
        })

    # 统计
    stat = {
        "S6(药物级HERB)": len(genes_s6), "S7(成分级HERB)": len(genes_s7),
        "S8(药物级TCMSP)": len(genes_s8), "S9(成分级TCMSP)": len(genes_s9),
        "S11(STP预测)": len(genes_s11),
        "并集(去重后)": len(all_genes),
        "S6∩S7": len(set(genes_s6) & set(genes_s7)),
        "S8∩S9": len(set(genes_s8) & set(genes_s9)),
        "S6∩S8": len(set(genes_s6) & set(genes_s8)),
        "S7∩S11": len(set(genes_s7) & set(genes_s11)),
        "S8∩S11": len(set(genes_s8) & set(genes_s11)),
        "数据库(S6∪S7∪S8∪S9)": len(set(genes_s6) | set(genes_s7) | set(genes_s8) | set(genes_s9)),
        "预测(S11)": len(genes_s11),
        "数据库∪预测": len(all_genes),
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
    cols0 = ["Gene symbol", "来源", "来源数", "涉及药物", "药物数", "UniProt ID"]
    ws0.append(cols0)
    for r in rows:
        ws0.append([r.get(c, "") for c in cols0])
    style_header(ws0, len(cols0))
    for i, c in enumerate(cols0, 1):
        ws0.column_dimensions[get_column_letter(i)].width = 16
    ws0.column_dimensions["A"].width = 18
    ws0.column_dimensions["D"].width = 60
    ws0.freeze_panes = "A2"

    ws1 = wb.create_sheet("来源统计")
    ws1.append(["统计项", "数量"])
    for k, v in stat.items():
        ws1.append([k, v])
    style_header(ws1, 2)
    ws1.column_dimensions["A"].width = 16
    ws1.column_dimensions["B"].width = 10

    wb.save(OUT)
    print(f"已写出 {OUT}: 去重后 Gene symbol {len(rows)} 个")


if __name__ == "__main__":
    main()
