# -*- coding: utf-8 -*-
"""
获取TCMSP成分.py  (增强版)
======================
从 TCMSP (https://www.tcmsp-e.com/) 获取 S1复方表.xlsx 中各药物的成分(分子)，
并补充 PubChem CID、InChIKey、CAS、TPSA、RBN、同义词 等字段，另存为单独文件。

输入文件:
    S1复方表.xlsx   —— 列1中文标准名、列2中文别名(顿号分隔)、列3拉丁名、列4英文名

输出文件:
    S2.TCMSP成分表.xlsx —— 药物汇总(含匹配名称) + 全部成分 + 每药一个工作表

流程:
    1. 读取 S1复方表.xlsx
    2. 获取 TCMSP 首页 token
    3. 按中文名搜索草药 (qs=herb_all_name)  -> 中文标准名优先, 未命中再按中文别名逐个搜索
    4. 按拉丁名搜索 (qsr=herb_en_name)      -> 页面内嵌分子列表(kendoGrid data); 中文名均未命中时用S1拉丁名直搜
    5. 对每个分子抓取 molecule.php?qn=<molecule_ID> 详情 -> PubChem CID/InChIKey/CAS/TPSA/RBN/同义词
    6. 汇总写出

用法:
    python 获取TCMSP成分.py
"""
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE = "https://www.tcmsp-e.com/"
HOME_URL = BASE + "tcmsp.php"
SEARCH_URL = BASE + "tcmspsearch.php"
MOLECULE_URL = BASE + "molecule.php"
S1_XLSX = "S1复方表.xlsx"
XLSX_OUT = "S2.TCMSP成分表.xlsx"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0 Safari/537.36",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}


def http_get(url, timeout=60, retries=3):
    last_err = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read()
            for enc in ("utf-8", "gb18030", "latin-1"):
                try:
                    return raw.decode(enc)
                except UnicodeDecodeError:
                    continue
            return raw.decode("utf-8", errors="replace")
        except Exception as ex:
            last_err = ex
            time.sleep(1.5 * (attempt + 1))
    raise last_err


def get_token():
    html = http_get(HOME_URL)
    m = re.search(r'name="token"\s+value="([0-9a-f]+)"', html)
    if not m:
        raise RuntimeError("无法从 TCMSP 首页获取 token")
    return m.group(1)


def extract_json_array(html, key):
    """提取 'key= [...]' 或 'key: [...]' 形式的 JSON 数组(括号配对)"""
    m = re.search(re.escape(key) + r"\s*[=:]?\s*(\[)", html)
    if not m:
        return None
    start = m.start(1)
    depth, in_str, esc = 0, False, False
    i = start
    while i < len(html):
        ch = html[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
        else:
            if ch == '"':
                in_str = True
            elif ch == "[":
                depth += 1
            elif ch == "]":
                depth -= 1
                if depth == 0:
                    break
        i += 1
    try:
        return json.loads(html[start:i + 1])
    except Exception:
        return None


def search_herb_by_cn(token, cn_name):
    url = SEARCH_URL + "?" + urllib.parse.urlencode(
        {"qs": "herb_all_name", "q": cn_name, "token": token})
    html = http_get(url)
    data = extract_json_array(html, "data")
    return data or []


def get_herb_molecules(token, latin_name):
    url = SEARCH_URL + "?" + urllib.parse.urlencode(
        {"qr": latin_name, "qsr": "herb_en_name", "token": token})
    html = http_get(url)
    data = extract_json_array(html, "data")
    return data or []


def parse_cas_cell(html):
    """解析 molecule.php 页面的 CAS 单元格"""
    m = re.search(r'<th class="left_header">CAS</th><td[^>]*>(.*?)</td>', html, re.S)
    if not m:
        return ""
    cell = m.group(1)
    cas_list = re.findall(r"\d{2,7}-\d{2}-\d{1}", cell)
    return "; ".join(cas_list)


def get_molecule_detail(mol_id_num):
    """抓取分子详情: 返回 PubChem CID / InChIKey / CAS / TPSA / RBN / 同义词"""
    url = f"{MOLECULE_URL}?qn={mol_id_num}"
    html = http_get(url)
    info = {}
    mol = extract_json_array(html, "var mol") or extract_json_array(html, "mol")
    if mol:
        m0 = mol[0]
        info["pubchem_cid"] = m0.get("pubchem_cid", "")
        info["inchikey"] = m0.get("inchikey", "")
        info["tpsa"] = m0.get("tpsa", "")
        info["rbn"] = m0.get("rbn", "")
        # 以详情页为准补齐 ADME
        for k in ("ob", "dl", "bbb", "caco2", "mw", "hdon", "hacc", "alogp",
                  "halflife", "FASA"):
            info[k] = m0.get(k, "")
    info["cas"] = parse_cas_cell(html)
    sname = extract_json_array(html, "var sname") or extract_json_array(html, "sname")
    if sname:
        info["synonyms"] = "; ".join(s.get("molecule_synonyms", "") for s in sname)
    else:
        info["synonyms"] = ""
    return info


def read_drugs_from_s1(path):
    """读取 S1复方表.xlsx: 中文标准名 / 中文别名(顿号分隔) / 拉丁名 / 英文名"""
    from openpyxl import load_workbook
    from collections import Counter

    wb = load_workbook(path, data_only=True)
    ws = wb.active
    drugs = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        cn = str(row[0]).strip()
        if not cn:
            continue
        alias_raw = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        latin = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
        en = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""
        aliases = [a for a in re.split(r"[、，,;；/／\s]+", alias_raw) if a]
        drugs.append({"标准名": cn, "别名": aliases, "拉丁名": latin, "英文名": en})
    cnt = Counter(d["标准名"] for d in drugs)
    for d in drugs:
        d["出现次数"] = cnt[d["标准名"]]
    return drugs


def find_herbs_by_names(token, names):
    """依次用多个中文名搜索 TCMSP 草药, 返回 (herbs, 命中的名称)"""
    for name in names:
        try:
            herbs = search_herb_by_cn(token, name)
        except Exception:
            herbs = []
        if herbs:
            return herbs, name
    return [], ""


def main():
    if not os.path.exists(S1_XLSX):
        print(f"错误: 找不到 {S1_XLSX}")
        sys.exit(1)

    drugs = read_drugs_from_s1(S1_XLSX)
    print(f"读取到 {len(drugs)} 味药物: " + "、".join(d["标准名"] for d in drugs))
    print("获取 TCMSP token ...")
    token = get_token()
    print(f"token = {token}")

    results, summary_rows, all_rows = [], [], []

    for idx, d in enumerate(drugs, 1):
        cn = d["标准名"]
        aliases = d["别名"]
        latin = d["拉丁名"]
        times = d["出现次数"]
        search_names = [cn] + aliases
        print(f"[{idx}/{len(drugs)}] 处理 {cn} ...", flush=True)

        try:
            herbs, matched_name = find_herbs_by_names(token, search_names)
        except Exception as ex:
            print(f"   搜索失败: {ex}")
            summary_rows.append({"药物": cn, "出现次数": times, "TCMSP中文名": "", "拼音": "",
                                 "拉丁名": "", "分子数": 0, "匹配名称": "",
                                 "状态": f"搜索出错: {ex}"})
            continue

        rows_for_herb, mols_by_id = [], {}
        if herbs:
            match_way = "标准名" if matched_name == cn else f"别名:{matched_name}"
            for h in herbs:
                h_cn = h.get("herb_cn_name", cn)
                h_py = h.get("herb_pinyin", "")
                h_en = h.get("herb_en_name", "")
                try:
                    mols = get_herb_molecules(token, h_en)
                except Exception as ex:
                    print(f"   拉丁名 {h_en} 分子获取失败: {ex}")
                    mols = []
                for m in mols:
                    rows_for_herb.append((cn, times, h_cn, h_py, h_en, m))
                    mid = m.get("molecule_ID", "")
                    if mid:
                        mols_by_id[mid] = m
                time.sleep(0.3)
        elif latin:
            # 中文名(标准名+别名)均未命中时, 直接用 S1 表中的拉丁名拉取分子
            match_way = f"拉丁名:{latin}"
            try:
                mols = get_herb_molecules(token, latin)
            except Exception as ex:
                print(f"   拉丁名 {latin} 分子获取失败: {ex}")
                mols = []
            for m in mols:
                rows_for_herb.append((cn, times, cn, "", latin, m))
                mid = m.get("molecule_ID", "")
                if mid:
                    mols_by_id[mid] = m
        else:
            match_way = "未找到"

        if not rows_for_herb:
            if herbs:
                status = "已匹配但无分子数据"
            elif latin:
                status = "拉丁名直搜无分子"
            else:
                status = "TCMSP 中未找到"
            summary_rows.append({"药物": cn, "出现次数": times, "TCMSP中文名": "", "拼音": "",
                                 "拉丁名": "", "分子数": 0, "匹配名称": match_way,
                                 "状态": status})
            continue

        # 并行抓取每个分子详情(按 molecule_ID 去重)
        print(f"   抓取 {len(mols_by_id)} 个分子详情 ...", flush=True)
        detail_map = {}
        with ThreadPoolExecutor(max_workers=8) as ex:
            futs = {ex.submit(get_molecule_detail, mid): mid for mid in mols_by_id}
            done_n = 0
            for fut in as_completed(futs):
                mid = futs[fut]
                try:
                    detail_map[mid] = fut.result()
                except Exception as e:
                    detail_map[mid] = {"pubchem_cid": "", "inchikey": "", "tpsa": "",
                                       "rbn": "", "cas": "", "synonyms": "",
                                       "detail_error": str(e)}
                done_n += 1
                if done_n % 100 == 0:
                    print(f"   分子详情进度 {done_n}/{len(mols_by_id)}", flush=True)

        for (cn_, times_, h_cn, h_py, h_en, m) in rows_for_herb:
            mid = m.get("molecule_ID", "")
            d = detail_map.get(mid, {})
            row = {
                "药物": cn_, "出现次数": times_, "TCMSP中文名": h_cn, "拼音": h_py,
                "拉丁名": h_en, "MOL_ID": m.get("MOL_ID", ""),
                "分子名": m.get("molecule_name", ""),
                "PubChem CID": d.get("pubchem_cid", ""),
                "CAS": d.get("cas", ""),
                "InChIKey": d.get("inchikey", ""),
                "MW": m.get("mw", ""), "OB(%)": m.get("ob", ""),
                "DL": m.get("dl", ""), "Caco-2": m.get("caco2", ""),
                "BBB": m.get("bbb", ""), "AlogP": m.get("alogp", ""),
                "HL": m.get("halflife", ""), "Hdon": m.get("hdon", ""),
                "Hacc": m.get("hacc", ""), "TPSA": d.get("tpsa", ""),
                "RBN": d.get("rbn", ""), "FASA": m.get("FASA", ""),
                "同义词": d.get("synonyms", ""),
            }
            results.append(row)
        all_rows.append((cn, herbs, rows_for_herb, detail_map))

        mol_ids = []
        for m in mols_by_id.values():
            if m.get("MOL_ID") and m["MOL_ID"] not in mol_ids:
                mol_ids.append(m["MOL_ID"])
        summary_rows.append({"药物": cn, "出现次数": times,
                             "TCMSP中文名": herbs[0].get("herb_cn_name", cn) if herbs else cn,
                             "拼音": herbs[0].get("herb_pinyin", "") if herbs else "",
                             "拉丁名": "; ".join(h.get("herb_en_name", "") for h in herbs) if herbs else latin,
                             "分子数": len(mol_ids), "匹配名称": match_way, "状态": "OK"})
        print(f"   {cn}: {len(mol_ids)} 个分子 (匹配: {match_way})", flush=True)


    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("未安装 openpyxl，跳过 xlsx 输出")
        return

    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="4472C4")
    head_font = Font(bold=True, color="FFFFFF")

    def style_header(ws, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=1, column=c)
            cell.fill = head_fill
            cell.font = head_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # 全部成分总表(供合并程序读取)
    ws_all = wb.create_sheet("全部成分")
    all_cols = ["药物", "出现次数", "TCMSP中文名", "拼音", "拉丁名", "MOL_ID", "分子名",
                "PubChem CID", "CAS", "InChIKey", "MW", "OB(%)", "DL", "Caco-2",
                "BBB", "AlogP", "HL", "Hdon", "Hacc", "TPSA", "RBN", "FASA", "同义词"]
    ws_all.append(all_cols)
    for r in results:
        ws_all.append([r.get(c, "") for c in all_cols])
    style_header(ws_all, len(all_cols))
    for i, c in enumerate(all_cols, 1):
        ws_all.column_dimensions[get_column_letter(i)].width = 14

    ws0 = wb.active
    ws0.title = "药物汇总"
    sum_cols = ["药物", "出现次数", "TCMSP中文名", "拼音", "拉丁名", "分子数", "匹配名称", "状态"]
    ws0.append(sum_cols)
    for s in summary_rows:
        ws0.append([s.get(c, "") for c in sum_cols])
    style_header(ws0, len(sum_cols))
    for i, c in enumerate(sum_cols, 1):
        ws0.column_dimensions[get_column_letter(i)].width = 16

    for cn, herbs, rows_for_herb, detail_map in all_rows:
        ws = wb.create_sheet(title=cn[:31])
        ws.append(["MOL_ID", "分子名", "PubChem CID", "CAS", "InChIKey", "MW",
                   "OB(%)", "DL", "Caco-2", "BBB", "AlogP", "HL", "Hdon", "Hacc",
                   "TPSA", "RBN", "FASA", "同义词", "拉丁名"])
        for (cn_, times_, h_cn, h_py, h_en, m) in rows_for_herb:
            mid = m.get("molecule_ID", "")
            d = detail_map.get(mid, {})
            ws.append([m.get("MOL_ID", ""), m.get("molecule_name", ""),
                       d.get("pubchem_cid", ""), d.get("cas", ""), d.get("inchikey", ""),
                       m.get("mw", ""), m.get("ob", ""), m.get("dl", ""),
                       m.get("caco2", ""), m.get("bbb", ""), m.get("alogp", ""),
                       m.get("halflife", ""), m.get("hdon", ""), m.get("hacc", ""),
                       d.get("tpsa", ""), d.get("rbn", ""), m.get("FASA", ""),
                       d.get("synonyms", ""), h_en])
        style_header(ws, 19)
        widths = [12, 40, 12, 26, 34, 10, 10, 8, 10, 8, 8, 8, 8, 8, 8, 8, 10, 40, 22]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(XLSX_OUT)
    print(f"已写出 {XLSX_OUT}")


if __name__ == "__main__":
    main()
