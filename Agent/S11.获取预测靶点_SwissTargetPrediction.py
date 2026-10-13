# -*- coding: utf-8 -*-
"""
S11.获取预测靶点_SwissTargetPrediction.py
=========================================
对"靶点总数<100"的药物，取其 S5.最终纳入.xlsx 中的成分，
用 SwissTargetPrediction (https://swisstargetprediction.ch/) 按 SMILES
做反向靶点预测，每个成分取前 100 个预测靶点（按概率排序）。

流程(每个唯一 SMILES, 串行提交):
    1. POST /predict.php (organism=Homo_sapiens, smiles=...) -> 阻塞等待计算完成
       (SwissTargetPrediction 一次只处理一个分子, 约 15~60 秒/个)
    2. 从返回页提取 location.replace("...result.php?job=NNN&organism=...")
    3. GET result.php -> 解析靶点表(前100行): Target, Common name(Gene symbol),
       Uniprot ID, ChEMBL ID, Target Class, Probability, Known actives

药物筛选:
    读取 S10.药物靶点统计.xlsx 的"靶点总数(去重)"列, 筛选 < 100 的药物。

输入:
    S10.药物靶点统计.xlsx (工作表: 药物靶点统计) —— 筛选靶点总数<100的药物
    S5.最终纳入.xlsx      (工作表: 最终纳入)   —— 获取这些药物的成分(SMILES)
输出:
    S11.预测靶点_SwissTargetPrediction.xlsx
    (工作表: 药物汇总 / 每药一个工作表[成分×前100靶点] / 全部靶点汇总)
缓存:
    cashe/stp_predictions.json (断点续跑)

用法:
    python S11.获取预测靶点_SwissTargetPrediction.py
    python S11.获取预测靶点_SwissTargetPrediction.py --limit 3      # 只预测前3个SMILES
    python S11.获取预测靶点_SwissTargetPrediction.py --drugs 水蛭   # 只处理指定药物
"""
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import http.cookiejar
from collections import defaultdict

S10_XLSX = "S10.药物靶点统计.xlsx"
S10_SHEET = "药物靶点统计"
S5_XLSX = "S5.最终纳入.xlsx"
S5_SHEET = "最终纳入"
XLSX_OUT = "S11.预测靶点_SwissTargetPrediction.xlsx"
CACHE = os.path.join("cashe", "stp_predictions.json")
TARGET_LIMIT = 100   # 每个成分取前 100 个预测靶点
DRUG_TARGET_THRESHOLD = 100

BASE = "https://swisstargetprediction.ch"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"

LIMIT = 0  # 0 = 全部; 可用命令行参数覆盖


def make_opener():
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    opener.addheaders = [("User-Agent", UA), ("Referer", BASE + "/")]
    return opener


def predict_one(opener, smiles, timeout=240):
    """提交一个 SMILES, 返回 job URL; INVALID=分子无效; None=失败"""
    data = urllib.parse.urlencode({
        "organism": "Homo_sapiens", "smiles": smiles, "Example": "", "ioi": "2"})
    req = urllib.request.Request(BASE + "/predict.php", data=data.encode(),
                                 headers={"Content-Type": "application/x-www-form-urlencoded"})
    for attempt in range(3):
        try:
            body = opener.open(req, timeout=timeout).read().decode("utf-8", errors="replace")
            m = re.search(r'location\.replace\("([^"]*result\.php[^"]*)"\)', body)
            if m:
                return m.group(1)
            if re.search(r"Query molecule is invalid", body):
                return "INVALID"
            print(f"   未找到结果跳转, 重试 {attempt+1}/3", flush=True)
        except Exception as ex:
            print(f"   predict 出错: {ex}, 重试 {attempt+1}/3", flush=True)
        time.sleep(3)
    return None


def parse_results(html):
    """解析 result.php 靶点表(按概率排序), 取前 TARGET_LIMIT 行"""
    rows = []
    m = re.search(r'<table id="resultTable".*?<tbody>(.*?)</tbody>', html, re.S)
    if not m:
        return rows
    tbody = m.group(1)
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", tbody, re.S):
        tds = re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S)
        if len(tds) < 6:
            continue
        def clean(cell):
            cell = re.sub(r"<[^>]+>", "", cell)
            return cell.strip()
        target = clean(tds[0])
        common = clean(tds[1])
        uniprot = clean(tds[2])
        chembl = clean(tds[3])
        tclass = clean(tds[4])
        prob = re.search(r"<span[^>]*>([\d.]+)</span>", tds[5])
        prob = prob.group(1) if prob else ""
        actives = re.sub(r"\s+", " ", clean(tds[6]) if len(tds) > 6 else "")
        rows.append({"Target": target, "Common name": common, "Uniprot ID": uniprot,
                     "ChEMBL ID": chembl, "Target Class": tclass,
                     "Probability": prob, "Known actives": actives})
        if len(rows) >= TARGET_LIMIT:
            break
    return rows


def fetch_result(opener, job_url):
    if job_url.startswith("/"):
        job_url = BASE + job_url
    req = urllib.request.Request(job_url, headers={"User-Agent": UA})
    for attempt in range(3):
        try:
            html = opener.open(req, timeout=90).read().decode("utf-8", errors="replace")
            return parse_results(html)
        except Exception as ex:
            print(f"   result 获取出错: {ex}, 重试 {attempt+1}/3", flush=True)
            time.sleep(3)
    return []


def load_cache():
    if os.path.exists(CACHE):
        try:
            with open(CACHE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_cache(cache):
    os.makedirs(os.path.dirname(CACHE), exist_ok=True)
    with open(CACHE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=1)


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


def select_drugs():
    """返回 S10 中靶点总数(去重)<100 的药物列表"""
    rows = read_xlsx(S10_XLSX, S10_SHEET)
    drugs = []
    for r in rows:
        drug = str(r.get("药物", "")).strip()
        total = r.get("靶点总数(去重)")
        if drug and drug != "合计" and isinstance(total, (int, float)) and total < DRUG_TARGET_THRESHOLD:
            drugs.append(drug)
    return drugs


def main():
    global LIMIT
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="只预测前N个唯一SMILES")
    ap.add_argument("--drugs", default="", help="只处理指定药物, 逗号分隔, 如: 水蛭,土鳖虫,赶黄草")
    args, _ = ap.parse_known_args()
    LIMIT = args.limit
    drug_filter = [d.strip() for d in args.drugs.split(",") if d.strip()] if args.drugs else []

    drugs = select_drugs()
    if drug_filter:
        drugs = [d for d in drugs if d in drug_filter]
    print(f"靶点总数<{DRUG_TARGET_THRESHOLD} 的药物: {'、'.join(drugs)}")
    if not drugs:
        print("没有需要补预测靶点的药物")
        return

    rows = read_xlsx(S5_XLSX, S5_SHEET)
    # 按药物筛选成分
    comp_rows = []
    for r in rows:
        drugs_of = [d.strip() for d in str(r.get("涉及药物", "")).split("、") if d.strip()]
        for d in drugs:
            if d in drugs_of:
                r2 = dict(r)
                r2["_药物"] = d
                comp_rows.append(r2)
    print(f"筛选到成分行: {len(comp_rows)} 条")

    # 有 SMILES 的行(优先 Canonical SMILES, 为空用异构SMILES)
    def get_smiles(r):
        s = str(r.get("Canonical SMILES", "") or "").strip()
        if s and s.upper() != "NA":
            return s
        s = str(r.get("异构SMILES", "") or "").strip()
        if s and s.upper() != "NA":
            return s
        return ""

    with_smi = [(r, get_smiles(r)) for r in comp_rows if get_smiles(r)]
    uniq = {}
    for r, s in with_smi:
        uniq.setdefault(s, []).append(r)
    print(f"有 SMILES 成分行 {len(with_smi)}, 唯一 SMILES {len(uniq)} 个")
    if LIMIT > 0:
        uniq = dict(list(uniq.items())[:LIMIT])
        print(f"本次只预测前 {LIMIT} 个 SMILES")

    cache = load_cache()
    todo = [s for s in uniq if s not in cache]
    print(f"待预测: {len(todo)} (缓存已完成 {len(uniq) - len(todo)})")

    opener = make_opener()
    try:
        opener.open(BASE + "/", timeout=60).read()
    except Exception as ex:
        print("首页会话建立失败:", ex)

    for k, s in enumerate(todo, 1):
        print(f"[{k}/{len(todo)}] 预测 {s[:50]} ...", flush=True)
        job_url = predict_one(opener, s)
        if job_url == "INVALID":
            cache[s] = {"status": "invalid", "targets": []}
            print("   SMILES 无效", flush=True)
        elif job_url:
            targets = fetch_result(opener, job_url)
            targets = targets[:TARGET_LIMIT]
            cache[s] = {"status": "ok", "job": job_url, "targets": targets}
            print(f"   获得预测靶点 {len(targets)} 个", flush=True)
        else:
            cache[s] = {"status": "failed", "targets": []}
            print("   预测失败", flush=True)
        save_cache(cache)
        time.sleep(1.5)

    # ---- 生成 xlsx ----
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    # 组装 药物-成分-靶点 记录
    records = []   # (药物, 成分ID, 成分名, SMILES, 排名, 靶点dict 或 None+备注)
    for s, recs in uniq.items():
        info = cache.get(s, {})
        status = info.get("status", "")
        targets = info.get("targets", [])
        for r in recs:
            drug = r.get("_药物", "")
            cid = str(r.get("成分ID", ""))
            iname = str(r.get("成分名", ""))
            if targets:
                for rank, t in enumerate(targets, 1):
                    records.append((drug, cid, iname, s, rank, t))
            else:
                note = "无SMILES" if not s else ("SMILES无效" if status == "invalid" else "预测失败/无结果")
                records.append((drug, cid, iname, s, "", {"_note": note}))

    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="4472C4")
    head_font = Font(bold=True, color="FFFFFF")

    def style_header(ws, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=1, column=c)
            cell.fill = head_fill
            cell.font = head_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # 药物汇总
    ws0 = wb.active
    ws0.title = "药物汇总"
    ws0.append(["药物", "成分数", "有SMILES成分数", "预测成功成分数", "去重预测靶点数"])
    by_drug = defaultdict(list)
    for s, recs in uniq.items():
        info = cache.get(s, {})
        for r in recs:
            by_drug[r.get("_药物", "")].append((s, info))
    for drug in drugs:
        items = by_drug.get(drug, [])
        n_total = sum(1 for r in comp_rows if r.get("_药物") == drug)
        n_smi = len(items)
        n_ok = sum(1 for _, info in items if info.get("status") == "ok")
        genes = set()
        for _, info in items:
            for t in info.get("targets", []):
                if t.get("Common name"):
                    genes.add(t["Common name"])
        ws0.append([drug, n_total, n_smi, n_ok, len(genes)])
    style_header(ws0, 5)
    for i, c in enumerate(["药物", "成分数", "有SMILES成分数", "预测成功成分数", "去重预测靶点数"], 1):
        ws0.column_dimensions[get_column_letter(i)].width = 16

    # 每药一个工作表
    per_drug = defaultdict(list)
    for rec in records:
        per_drug[rec[0]].append(rec)
    for drug in drugs:
        recs = per_drug.get(drug, [])
        title = drug[:31]
        ws = wb.create_sheet(title=title)
        ws.append(["成分ID", "成分名", "SMILES", "排名", "Gene symbol", "UniProt ID",
                   "ChEMBL ID", "靶点名", "靶点类别", "概率", "Known actives", "备注"])
        for (drug_, cid, iname, s, rank, t) in recs:
            ws.append([cid, iname, s, rank, t.get("Common name", ""), t.get("Uniprot ID", ""),
                       t.get("ChEMBL ID", ""), t.get("Target", ""), t.get("Target Class", ""),
                       t.get("Probability", ""), t.get("Known actives", ""), t.get("_note", "")])
        style_header(ws, 12)
        for i, w in enumerate([14, 40, 50, 8, 14, 12, 14, 45, 18, 12, 14, 16], 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    # 全部靶点汇总
    target_drugs = defaultdict(set)
    target_info = {}
    for (drug, cid, iname, s, rank, t) in records:
        g = t.get("Common name", "")
        if not g:
            continue
        target_drugs[g].add(drug)
        if g not in target_info:
            target_info[g] = t
    ws_all = wb.create_sheet("全部靶点汇总")
    ws_all.append(["Gene symbol", "UniProt ID", "ChEMBL ID", "靶点名", "靶点类别",
                   "最高概率", "涉及药物数", "涉及药物"])
    for g in sorted(target_drugs.keys()):
        t = target_info[g]
        dl = sorted(target_drugs[g])
        ws_all.append([g, t.get("Uniprot ID", ""), t.get("ChEMBL ID", ""),
                       t.get("Target", ""), t.get("Target Class", ""),
                       t.get("Probability", ""), len(dl), "; ".join(dl)])
    style_header(ws_all, 8)
    for i, w in enumerate([14, 12, 14, 45, 18, 12, 12, 60], 1):
        ws_all.column_dimensions[get_column_letter(i)].width = w

    wb.save(XLSX_OUT)
    n_ok = sum(1 for v in cache.values() if v.get("status") == "ok")
    print(f"已写出 {XLSX_OUT}: 预测成功 {n_ok} / {len(cache)} 个SMILES")


if __name__ == "__main__":
    main()
