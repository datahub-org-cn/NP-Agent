# -*- coding: utf-8 -*-
"""
S4.去除假阳性数据.py
====================
对 S3 筛选后的"纳入成分"做假阳性排查(溯源验证):

    第1步 标识符完整性: 每行成分必须至少含有以下字段之一
        PubChem CID、PubChem 名称、CAS、分子式、Canonical SMILES、
        异构SMILES、InChI、InChIKey、IUPAC名称
        一个都没有 -> 判定"无标识符"，剔除。
        (注: 分子式不唯一, 不能单独作为溯源依据; 仅含分子式 -> "仅有分子式(无法溯源)", 剔除)

    第2步 数据库溯源: 用已有标识符到 PubChem (PUG REST) 查询, 按优先级分层:
        InChIKey > PubChem CID > InChI > Canonical SMILES > 异构SMILES
        > PubChem名称 / IUPAC名称 / 成分名(分子名) / 同义词 > CAS
        查到对应 PubChem 化合物 -> "找到", 保留;
        所有标识符均查不到       -> "未找到", 剔除(假阳性);
        网络/服务异常            -> "溯源异常(暂保留)", 保留待复核。
    说明: 按优先级"分层查询", 只解析各行实际需要的标识符, 避免无效请求。

输入:
    S3.HERB成分表-OB-DL--Lipinski.xlsx  (工作表: 纳入成分)
    S3.TCMSP成分表-DL-OB-Lipinski.xlsx  (工作表: 纳入成分)

输出:
    S4.HERB成分表-纳入.xlsx   (工作表: 药物汇总 / 溯源明细 / 纳入成分)
    S4.TCMSP成分表-纳入.xlsx  (工作表: 药物汇总 / 溯源明细 / 纳入成分)
    cashe/pubchem_cache.json  (PubChem 查询缓存, 重复运行不重复请求)

用法:
    python S4.去除假阳性数据.py
"""
import json
import os
import re
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed

PUG = "https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound"
CACHE_FILE = "cashe/pubchem_cache.json"
MAX_WORKERS = 4
RATE_PER_SEC = 4.0
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

DB_CONFIG = {
    "HERB": {
        "INPUT": "S3.HERB成分表-OB-DL--Lipinski.xlsx",
        "OUT": "S4.HERB成分表-纳入.xlsx",
        "PRESENCE": ["PubChem CID", "PubChem 名称", "CAS", "分子式", "Canonical SMILES",
                     "异构SMILES", "InChI", "InChIKey", "IUPAC名称", "成分名"],
        # (查询类型, 列名), 按优先级排列
        "TRACE": [
            ("inchikey", "InChIKey"),
            ("cid", "PubChem CID"),
            ("inchi", "InChI"),
            ("smiles", "Canonical SMILES"),
            ("smiles", "异构SMILES"),
            ("name", "PubChem 名称"),
            ("name", "IUPAC名称"),
            ("name", "成分名"),
            ("name", "CAS"),
        ],
    },
    "TCMSP": {
        "INPUT": "S3.TCMSP成分表-DL-OB-Lipinski.xlsx",
        "OUT": "S4.TCMSP成分表-纳入.xlsx",
        "PRESENCE": ["PubChem CID", "CAS", "InChIKey", "分子名", "同义词"],
        "TRACE": [
            ("inchikey", "InChIKey"),
            ("cid", "PubChem CID"),
            ("name", "分子名"),
            ("name", "同义词"),
            ("name", "CAS"),
        ],
    },
}

EXTRA_COLS = ["标识符字段数", "溯源结果", "溯源依据", "溯源CID", "溯源查询明细"]
KEEP_RESULTS = ("找到", "溯源异常(暂保留)")


def read_xlsx(path, sheet):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    out = []
    for r in rows[1:]:
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        d = {}
        for i, name in enumerate(header):
            if i < len(r):
                d[name] = r[i]
        out.append(d)
    return out


def nonempty(v):
    if v is None:
        return False
    s = str(v).strip()
    return bool(s) and s.upper() != "NA"


# ---------------- PubChem 查询 ----------------
_lock = threading.Lock()
_last_req = [0.0]


def throttle():
    with _lock:
        now = time.time()
        wait = _last_req[0] + 1.0 / RATE_PER_SEC - now
        if wait > 0:
            time.sleep(wait)
        _last_req[0] = time.time()


def pubchem_lookup(field, value):
    """返回命中的 PubChem CID 列表; 查不到返回 None; 网络/服务异常返回 'ERR'"""
    q = urllib.parse.quote(value, safe="")
    data = None
    if field == "cid":
        url = f"{PUG}/cid/{q}/cids/JSON"
    elif field == "inchikey":
        url = f"{PUG}/inchikey/{q}/cids/JSON"
    elif field == "inchi":
        # PUG REST 的 inchi 路径含 "/", GET 解析失败, 改用 POST
        url = f"{PUG}/inchi/cids/JSON"
        data = urllib.parse.urlencode({"inchi": value}).encode("utf-8")
    elif field == "smiles":
        url = f"{PUG}/smiles/{q}/cids/JSON"
    elif field == "name":
        url = f"{PUG}/name/{q}/cids/JSON"
    else:
        return None

    for attempt in range(4):
        throttle()
        try:
            headers = dict(UA)
            if data is not None:
                headers["Content-Type"] = "application/x-www-form-urlencoded"
            req = urllib.request.Request(url, data=data, headers=headers)
            with urllib.request.urlopen(req, timeout=45) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
            cids = payload.get("IdentifierList", {}).get("CID") or []
            return cids if cids else None
        except urllib.error.HTTPError as e:
            if e.code in (400, 404):
                return None
            if e.code == 429:
                time.sleep(3 * (attempt + 1))
                continue
            time.sleep(2 * (attempt + 1))
        except Exception:
            time.sleep(2 * (attempt + 1))
    return "ERR"


# ---------------- 缓存 ----------------
def load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, encoding="utf-8") as f:
                raw = json.load(f)
            return {tuple(k.split("\u0001", 1)): v for k, v in raw.items()}
        except Exception:
            return {}
    return {}


def save_cache(cache):
    os.makedirs(os.path.dirname(CACHE_FILE) or ".", exist_ok=True)
    raw = {"\u0001".join(k): v for k, v in cache.items()}
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(raw, f, ensure_ascii=False, indent=1)


# ---------------- 标识符收集与判定 ----------------
def collect_candidates(row, cfg):
    seen, out = set(), []
    for field, key in cfg["TRACE"]:
        raw = row.get(key)
        if not nonempty(raw):
            continue
        s = str(raw).strip()
        if key == "同义词":
            parts = [p.strip() for p in re.split(r"[;；]+", s) if p.strip()][:3]
        else:
            parts = [s]
        for p in parts:
            if not p:
                continue
            k = (field, p)
            if k not in seen:
                seen.add(k)
                out.append(k)
    return out


def build_detail(cands, upto, cache):
    parts = []
    for i in range(upto + 1):
        field, value = cands[i]
        res = cache.get((field, value))
        short = value if len(value) <= 40 else value[:37] + "..."
        tag = "找到" if isinstance(res, list) and res else ("异常" if res == "ERR" else "未找到")
        parts.append(f"{field}:{short}→{tag}")
    return "; ".join(parts)


# ---------------- 输出 ----------------
def write_xlsx(db_name, cfg, rows, orig_cols):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    kept = [r for r in rows if r["溯源结果"] in KEEP_RESULTS]
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
    ws0.append(["药物", "S3纳入", "溯源找到", "溯源异常", "剔除"])
    drugs = []
    for r in rows:
        if r["药物"] not in drugs:
            drugs.append(r["药物"])
    cnt_ok = Counter(r["药物"] for r in rows if r["溯源结果"] == "找到")
    cnt_er = Counter(r["药物"] for r in rows if r["溯源结果"] == "溯源异常(暂保留)")
    cnt_bad = Counter(r["药物"] for r in rows if r["溯源结果"] not in KEEP_RESULTS)
    for drug in drugs:
        ws0.append([drug, sum(1 for r in rows if r["药物"] == drug),
                    cnt_ok.get(drug, 0), cnt_er.get(drug, 0), cnt_bad.get(drug, 0)])
    ws0.append(["合计", len(rows),
                len([r for r in rows if r["溯源结果"] == "找到"]),
                len([r for r in rows if r["溯源结果"] == "溯源异常(暂保留)"]),
                len(rows) - len(kept)])
    style_header(ws0, 5)
    for i, c in enumerate(["药物", "S3纳入", "溯源找到", "溯源异常", "剔除"], 1):
        ws0.column_dimensions[get_column_letter(i)].width = 13

    cols = orig_cols + EXTRA_COLS
    wide_kw = ("SMILES", "InChI", "Morgan", "同义词", "别名", "IUPAC", "分子式")

    def make_sheet(title, data_rows):
        ws = wb.create_sheet(title)
        ws.append(cols)
        for r in data_rows:
            ws.append([r.get(c, "") for c in cols])
        style_header(ws, len(cols))
        for i, c in enumerate(cols, 1):
            if c in ("InChIKey",):
                ws.column_dimensions[get_column_letter(i)].width = 34
            elif c in ("CAS",):
                ws.column_dimensions[get_column_letter(i)].width = 18
            elif any(k in c for k in wide_kw):
                ws.column_dimensions[get_column_letter(i)].width = 30
            else:
                ws.column_dimensions[get_column_letter(i)].width = 12
        ws.freeze_panes = "A2"
        return ws

    make_sheet("溯源明细", rows)
    make_sheet("纳入成分", kept)

    wb.save(cfg["OUT"])
    print(f"已写出 {cfg['OUT']}: 保留 {len(kept)} / {len(rows)}")


def main():
    all_rows, per_db = [], {}
    for name, cfg in DB_CONFIG.items():
        if not os.path.exists(cfg["INPUT"]):
            print(f"错误: 找不到 {cfg['INPUT']} (跳过 {name})")
            continue
        rows = read_xlsx(cfg["INPUT"], "纳入成分")
        print(f"[{name}] 读取 {cfg['INPUT']} 工作表[纳入成分]: {len(rows)} 行")
        orig_cols = list(rows[0].keys()) if rows else []
        per_db[name] = (cfg, rows, orig_cols)
        for r in rows:
            all_rows.append((name, cfg, r))
    if not all_rows:
        print("无输入数据")
        sys.exit(1)

    # 每行状态: [cfg, row, 候选列表, 当前候选序号, 判定]
    states = []
    for name, cfg, r in all_rows:
        cands = collect_candidates(r, cfg)
        states.append([cfg, r, cands, 0, None])

    cache = load_cache()
    total_resolved = 0

    # 按优先级分层查询: 每轮只解析各行"当前最优先"且未缓存的标识符
    while True:
        needed = {}
        any_unresolved = False
        for st in states:
            cfg, r, cands, idx, verdict = st
            if verdict is not None:
                continue
            if not cands:
                st[4] = "nocand"
                continue
            if idx >= len(cands):
                st[4] = "done"
                continue
            any_unresolved = True
            key = cands[idx]
            if key not in cache:
                needed.setdefault(key, None)
        if not any_unresolved:
            break
        if needed:
            todo = list(needed)
            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
                futs = {ex.submit(pubchem_lookup, f, v): (f, v) for (f, v) in todo}
                for fut in as_completed(futs):
                    key = futs[fut]
                    try:
                        cache[key] = fut.result()
                    except Exception:
                        cache[key] = "ERR"
                    total_resolved += 1
                    if total_resolved % 50 == 0:
                        print(f"   已解析标识符 {total_resolved} 个", flush=True)
            save_cache(cache)
        # 依据本轮结果推进各行
        for st in states:
            cfg, r, cands, idx, verdict = st
            if verdict is not None:
                continue
            if not cands or idx >= len(cands):
                continue
            key = cands[idx]
            res = cache.get(key)
            if isinstance(res, list) and res:
                st[4] = ("found", key, str(res[0]))
            elif res == "ERR":
                st[4] = ("err", key)
            else:
                st[4] = None
                st[3] = idx + 1   # 前进到下一优先级的候选

    # 生成每行判定
    for cfg, r, cands, idx, verdict in states:
        present = sum(1 for k in cfg["PRESENCE"] if nonempty(r.get(k)))
        r["标识符字段数"] = present
        if verdict == "nocand":
            r["溯源结果"] = "仅有分子式(无法溯源)" if present else "无标识符"
            r["溯源依据"] = ""
            r["溯源CID"] = ""
            r["溯源查询明细"] = ""
        elif verdict == "done":
            r["溯源结果"] = "未找到"
            r["溯源依据"] = ""
            r["溯源CID"] = ""
            r["溯源查询明细"] = build_detail(cands, len(cands) - 1, cache) if cands else ""
        elif verdict[0] == "found":
            key, cid = verdict[1], verdict[2]
            r["溯源结果"] = "找到"
            r["溯源依据"] = f"{key[0]}={key[1]}"
            r["溯源CID"] = cid
            r["溯源查询明细"] = build_detail(cands, idx, cache)
        else:  # "err"
            key = verdict[1]
            r["溯源结果"] = "溯源异常(暂保留)"
            r["溯源依据"] = f"{key[0]}={key[1]}"
            r["溯源CID"] = ""
            r["溯源查询明细"] = build_detail(cands, idx, cache)

    stat = Counter(r["溯源结果"] for _, r, _, _, _ in states)
    print("溯源统计:", dict(stat))
    print(f"实际解析标识符: {total_resolved} 个 (含缓存命中)")

    for name, (cfg, rows, orig_cols) in per_db.items():
        write_xlsx(name, cfg, rows, orig_cols)
    print("=== S4 去除假阳性完成 ===")


if __name__ == "__main__":
    main()
