# -*- coding: utf-8 -*-
"""
S17.调用Cytoscape绘制中药复方调控网络.py
=========================================
通过 Cytoscape cyREST (http://127.0.0.1:1234) 绘制"中药复方调控网络"：
    药物 -> 成分 -> 核心靶点 -> 疾病

数据:
    药物/成分      : S1复方表.xlsx + S5.最终纳入.xlsx (涉及药物)
    核心靶点        : S16.Venn图-药物vs疾病.xlsx (交集基因, 118个)
    成分-靶点映射   : 双来源共同参与后续计算
                      1) 数据库靶点: S7.成分级靶点_HERB / S9.成分级靶点_TCMSP(补全GeneSymbol),
                                    经 S4 的 PubChem CID 关联到 S5 成分
                      2) 预测靶点  : S11.预测靶点_SwissTargetPrediction.xlsx (按S5成分ID)
    边"证据来源"属性: 数据库 / 预测 / 数据库;预测

流程:
    1. 组装节点/边(药物-成分 / 成分-靶点 / 靶点-疾病)
    2. cyREST 创建网络(POST /networks)
    3. 创建样式(POST /styles, title格式) + vizmap apply
    4. 固定布局：疾病居中、靶点多层同心圆、八味中药及其成分各成圆形簇
    5. 通过 Cytoscape 视图接口导出高分辨率 PNG 与 SVG，并保存 CYS 会话

前置: 已安装并启动 Cytoscape (cyREST 端口1234)
用法: python S17.调用Cytoscape绘制中药复方调控网络.py
"""
import json
import math
import os
import sys
import urllib.request
from collections import defaultdict
from datetime import datetime
from urllib.parse import quote

CYTO = "http://127.0.0.1:1234/v1"
PNG_OUT = "S17.中药复方调控网络-模板布局.png"
SVG_OUT = "S17.中药复方调控网络-模板布局.svg"
SESSION_OUT = "S17.中药复方调控网络-模板布局.cys"
XLSX_OUT = "S17.中药复方调控网络数据.xlsx"

S1 = "S1复方表.xlsx"
S5 = "S5.最终纳入.xlsx"
S4H = "S4.HERB成分表-纳入.xlsx"
S4T = "S4.TCMSP成分表-纳入.xlsx"
S7H = "S7.成分级靶点_HERB.xlsx"
S9T = "S9.成分级靶点_TCMSP-补全GeneSymbol.xlsx"
S16 = "S16.Venn图-药物vs疾病.xlsx"
S11 = "S11.预测靶点_SwissTargetPrediction.xlsx"

DISEASE_NODE = "疾病:DepressiveDisorder"
DISEASE_LABEL = "抑郁障碍"

# 按模板.jpg的顺时针花瓣式布局排列八味药。
DRUG_ORDER = ["柴胡", "当归", "白芍", "白术", "茯苓", "甘草", "薄荷", "生姜"]
GROUP_COLORS = {
    "柴胡": "#7569A7",  # 紫
    "当归": "#E31A1C",  # 红
    "白芍": "#E6B800",  # 黄
    "白术": "#263B8C",  # 深蓝
    "茯苓": "#D99AA6",  # 粉
    "甘草": "#32A7B5",  # 青
    "薄荷": "#E8862D",  # 橙
    "生姜": "#AA5A9E",  # 紫红
    "靶点": "#86B93F",  # 模板中央绿色
    "疾病": "#F39C34",
}


def rest(method, path, payload=None, timeout=120):
    url = CYTO + path
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method,
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        raw = r.read()
    if not raw:
        return None
    try:
        return json.loads(raw.decode("utf-8"))
    except Exception:
        return raw


def download(path, output_path, timeout=180):
    """从 Cytoscape cyREST 下载二进制导出结果。"""
    with urllib.request.urlopen(CYTO + path, timeout=timeout) as r:
        data = r.read()
    with open(output_path, "wb") as f:
        f.write(data)
    return len(data)


def read(path, sheet):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(h) if h is not None else "" for h in rows[0]]
    out = []
    for r in rows[1:]:
        out.append({hdr[i]: r[i] for i in range(len(hdr))})
    wb.close()
    return out


def build_network():
    # 核心靶点
    core = set()
    for d in read(S16, "交集基因(核心靶点)"):
        g = str(d.get("Gene symbol") or "").strip()
        if g:
            core.add(g.upper())

    # S5 成分 (CID -> 记录)
    s5 = {}
    for d in read(S5, "最终纳入"):
        cid = str(d.get("PubChem CID") or "").strip()
        if cid:
            s5[cid] = d

    # S4: CID -> HBIN / MOL_ID
    cid_hbin, cid_mol = defaultdict(set), defaultdict(set)
    for d in read(S4H, "纳入成分"):
        cid = str(d.get("PubChem CID") or "").strip()
        iid = str(d.get("成分ID") or "").strip()
        if cid and iid:
            cid_hbin[cid].add(iid)
    for d in read(S4T, "纳入成分"):
        cid = str(d.get("PubChem CID") or "").strip()
        mol = str(d.get("MOL_ID") or "").strip()
        if cid and mol:
            cid_mol[cid].add(mol)

    # S7/S9: 成分ID/MOL_ID -> 基因
    hbin_genes, mol_genes = defaultdict(set), defaultdict(set)
    for d in read(S7H, "靶点明细"):
        iid = str(d.get("成分ID") or "").strip()
        g = str(d.get("Gene symbol") or "").strip()
        if iid and g:
            hbin_genes[iid].add(g.upper())
    for d in read(S9T, "靶点明细"):
        mol = str(d.get("MOL_ID") or "").strip()
        g = str(d.get("Gene symbol") or "").strip()
        if mol and g:
            mol_genes[mol].add(g.upper())

    # S11 STP 预测靶点: S5成分ID(Cxxxx) -> 基因集
    stp_genes = defaultdict(set)
    from openpyxl import load_workbook as _lw
    _wb11 = _lw(S11, read_only=True, data_only=True)
    for sheet in _wb11.sheetnames:
        if sheet in ("药物汇总", "全部靶点汇总"):
            continue
        _ws = _wb11[sheet]
        _rows = list(_ws.iter_rows(values_only=True))
        _hdr = [str(h) if h is not None else "" for h in _rows[0]]
        if "成分ID" not in _hdr or "Gene symbol" not in _hdr:
            continue
        _ci, _gi = _hdr.index("成分ID"), _hdr.index("Gene symbol")
        for r in _rows[1:]:
            cid5 = str(r[_ci] or "").strip()
            g = str(r[_gi] or "").strip().upper()
            if cid5 and g and g != "N/A":
                stp_genes[cid5].add(g)
    _wb11.close()
    # S5 成分ID(Cxxxx) -> PubChem CID, 统一键控
    id2cid = {str(d.get("成分ID") or "").strip(): cid for cid, d in s5.items()}
    stp_genes = {id2cid.get(k, k): v for k, v in stp_genes.items()}
    print(f"S11 预测靶点: {len(stp_genes)} 个S5成分有STP预测")

    # 成分 -> 核心靶点(数据库S7/S9 + 预测S11, 记录证据来源)
    comp_core = {}      # cid -> {gene: "数据库"/"预测"/"数据库;预测"}
    for cid, d in s5.items():
        genes = {}
        db_hit = set()
        for h in cid_hbin.get(cid, []):
            db_hit |= hbin_genes.get(h, set())
        for m in cid_mol.get(cid, []):
            db_hit |= mol_genes.get(m, set())
        db_hit &= core
        for g in db_hit:
            genes[g] = "数据库"
        for g in (stp_genes.get(cid, set()) & core):
            genes[g] = "预测" if g not in genes else "数据库;预测"
        if genes:
            comp_core[cid] = genes

    # 节点
    nodes, edges = [], []
    node_names, node_labels = set(), {}
    def add_node(name, label, ntype, group="", full_name=""):
        if name in node_names:
            return
        node_names.add(name)
        node_labels[name] = label
        nodes.append({"data": {"id": name, "name": name, "label": label,
                               "type": ntype, "group": group or ntype,
                               "full_name": full_name or label}})
    def add_edge(s, t, etype):
        edges.append({"data": {"source": s, "target": t, "type": etype}})

    # 药物节点(仅保留有成分边的)
    drug_comp = defaultdict(set)
    for cid, hit in comp_core.items():
        for x in str(s5[cid].get("涉及药物") or "").split("、"):
            x = x.strip()
            if x:
                drug_comp[x].add(cid)
    ordered_drugs = [x for x in DRUG_ORDER if x in drug_comp]
    ordered_drugs += sorted(set(drug_comp) - set(ordered_drugs))
    for drug in ordered_drugs:
        add_node(drug, drug, "药物", drug)

    for cid, hit in comp_core.items():
        name = f"comp_{cid}"
        linked_drugs = [d for d in ordered_drugs if cid in drug_comp[d]]
        primary_drug = linked_drugs[0] if linked_drugs else "成分"
        short_label = str(s5[cid].get("成分ID") or cid).strip()
        full_name = str(s5[cid].get("成分名") or cid).strip()
        add_node(name, short_label, "成分", primary_drug, full_name)
        for drug in drug_comp:
            if cid in drug_comp[drug]:
                add_edge(drug, name, "药物-成分")
        for g, ev in hit.items():
            add_node(g, g, "靶点", "靶点")
            add_edge(name, g, "成分-靶点")
            edges[-1]["data"]["evidence"] = ev
    add_node(DISEASE_NODE, DISEASE_LABEL, "疾病", "疾病")
    for g in sorted({g for hit in comp_core.values() for g in hit}):
        add_edge(g, DISEASE_NODE, "靶点-疾病")

    ntype_cnt = defaultdict(int)
    for n in nodes:
        ntype_cnt[n["data"]["type"]] += 1
    print("节点:", dict(ntype_cnt), "总数", len(nodes))
    print("边:", len(edges), "| 药物-成分:", sum(1 for e in edges if e["data"]["type"]=="药物-成分"),
          "成分-靶点:", sum(1 for e in edges if e["data"]["type"]=="成分-靶点"),
          "靶点-疾病:", sum(1 for e in edges if e["data"]["type"]=="靶点-疾病"))
    return nodes, edges


def template_positions(nodes):
    """生成“疾病居中 + 靶点同心圆 + 外围八个药物-成分圆”的固定坐标。"""
    positions = {}

    # 疾病节点固定在全图正中心。
    positions[DISEASE_NODE] = (0.0, 0.0)

    # 靶点围绕疾病排成多层同心圆；每层容量随周长增加。
    targets = sorted((n for n in nodes if n["data"]["type"] == "靶点"),
                     key=lambda n: n["data"]["label"])
    target_index = 0
    ring_index = 0
    while target_index < len(targets):
        radius = 125.0 + ring_index * 54.0
        capacity = max(16, int(2.0 * math.pi * radius / 42.0))
        batch = targets[target_index:target_index + capacity]
        phase = (ring_index % 2) * math.pi / max(1, len(batch))
        for j, node in enumerate(batch):
            angle = phase + 2.0 * math.pi * j / len(batch)
            positions[node["data"]["id"]] = (
                radius * math.cos(angle),
                radius * math.sin(angle),
            )
        target_index += len(batch)
        ring_index += 1

    groups = defaultdict(list)
    for node in nodes:
        if node["data"]["type"] == "成分":
            groups[node["data"]["group"]].append(node)
    for members in groups.values():
        members.sort(key=lambda n: n["data"]["label"])

    present_drugs = [d for d in DRUG_ORDER if any(
        n["data"]["type"] == "药物" and n["data"]["id"] == d for n in nodes)]
    present_drugs += sorted(set(groups) - set(present_drugs))
    total_groups = max(1, len(present_drugs))

    # 八个药物圆均匀分布在核心靶点圆外围；中药位于各自圆心。
    start_deg = -157.5
    step_deg = 360.0 / total_groups
    for gi, drug in enumerate(present_drugs):
        center_deg = start_deg + gi * step_deg
        center_rad = math.radians(center_deg)
        center_x = 1130.0 * math.cos(center_rad)
        center_y = 970.0 * math.sin(center_rad)
        positions[drug] = (center_x, center_y)
        members = groups.get(drug, [])
        member_index = 0
        local_ring = 0
        while member_index < len(members):
            local_radius = 82.0 + local_ring * 54.0
            capacity = max(12, int(2.0 * math.pi * local_radius / 39.0))
            batch = members[member_index:member_index + capacity]
            phase = center_rad + (local_ring % 2) * math.pi / max(1, len(batch))
            for j, node in enumerate(batch):
                angle = phase + 2.0 * math.pi * j / len(batch)
                positions[node["data"]["id"]] = (
                    center_x + local_radius * math.cos(angle),
                    center_y + local_radius * math.sin(angle),
                )
            member_index += len(batch)
            local_ring += 1
    return positions


def apply_positions(suid, view_id, nodes, positions):
    """一次批量写入全部节点坐标，坐标作为 bypass 固定在 Cytoscape 视图中。"""
    network_json = rest("GET", f"/networks/{suid}")
    cyt_nodes = network_json.get("elements", {}).get("nodes", [])
    id_to_suid = {}
    for node in cyt_nodes:
        data = node.get("data", {})
        # Cytoscape 将导入的字符串 id 保存为 id_original/name，id 会改成内部 SUID。
        node_id = str(data.get("id_original") or data.get("name") or data.get("id") or "")
        node_suid = data.get("SUID")
        if node_id and node_suid is not None:
            id_to_suid[node_id] = int(node_suid)
    payload = []
    for node in nodes:
        node_id = node["data"]["id"]
        if node_id not in id_to_suid or node_id not in positions:
            continue
        x, y = positions[node_id]
        payload.append({
            "SUID": id_to_suid[node_id],
            "view": [
                {"visualProperty": "NODE_X_LOCATION", "value": float(x)},
                {"visualProperty": "NODE_Y_LOCATION", "value": float(y)},
            ],
        })
    rest("PUT", f"/networks/{suid}/views/{view_id}/nodes?bypass=true", payload, timeout=180)
    print(f"模板坐标已写入: {len(payload)}/{len(nodes)} 个节点")


def main():
    if not all(os.path.exists(p) for p in (S1, S5, S4H, S4T, S7H, S9T, S16)):
        print("错误: 缺少输入文件")
        sys.exit(1)

    # 先构建并保存网络数据，确保 Cytoscape 未启动时仍可完成数据分析。
    nodes, edges = build_network()
    if not nodes:
        print("无网络数据")
        sys.exit(1)
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "网络数据"
    ws.append(["节点/边", "来源", "目标", "类型", "标签", "分组", "完整名称", "证据来源"])
    for n in nodes:
        ws.append(["节点", n["data"]["id"], "", n["data"]["type"], n["data"]["label"],
                   n["data"].get("group", ""), n["data"].get("full_name", ""), ""])
    for e in edges:
        ws.append(["边", e["data"]["source"], e["data"]["target"], e["data"]["type"], "",
                   "", "", e["data"].get("evidence", "")])
    wb.save(XLSX_OUT)
    print(f"已写出 {XLSX_OUT} ({len(nodes)}节点/{len(edges)}边)")

    try:
        info = rest("GET", "/")
        print("cyREST 版本:", info.get("apiVersion"))
    except Exception as e:
        print("无法连接 Cytoscape cyREST:", e)
        print("已保留网络数据；启动 Cytoscape 后可重跑以补充图片")
        return

    print("创建网络 ...")
    net = rest("POST", "/networks", {
        "data": {"name": "中药复方调控网络"},
        "elements": {"nodes": nodes, "edges": edges},
    })
    if isinstance(net, dict) and net.get("errors"):
        print("建网失败:", net["errors"])
        sys.exit(1)
    suid = net["networkSUID"]
    print("网络 SUID:", suid)

    # 样式：颜色按中药分组，形状/大小按节点类型；整体贴近模板.jpg。
    color_map = [{"key": key, "value": value} for key, value in GROUP_COLORS.items()]
    style_payload = {
        "title": "FormulaTemplateStyle_" + datetime.now().strftime("%Y%m%d_%H%M%S"),
        "defaults": [
            {"visualProperty": "NETWORK_BACKGROUND_PAINT", "value": "#FFFFFF"},
            {"visualProperty": "NODE_FILL_COLOR", "value": "#86B93F"},
            {"visualProperty": "NODE_LABEL_COLOR", "value": "#000000"},
            {"visualProperty": "NODE_BORDER_WIDTH", "value": 0.6},
            {"visualProperty": "NODE_BORDER_PAINT", "value": "#5F666B"},
            {"visualProperty": "EDGE_STROKE_UNSELECTED_PAINT", "value": "#7B8185"},
            {"visualProperty": "EDGE_TARGET_ARROW_SHAPE", "value": "NONE"},
            {"visualProperty": "EDGE_TRANSPARENCY", "value": 78},
            {"visualProperty": "EDGE_WIDTH", "value": 0.55},
            {"visualProperty": "NODE_LABEL_FONT_SIZE", "value": 7},
            {"visualProperty": "NODE_LABEL_POSITION", "value": "C,C,c,0.00,0.00"},
        ],
        "mappings": [
            {"mappingType": "passthrough", "mappingColumn": "label", "mappingColumnType": "String",
             "visualProperty": "NODE_LABEL"},
            {"mappingType": "discrete", "mappingColumn": "group", "mappingColumnType": "String",
             "visualProperty": "NODE_FILL_COLOR", "map": color_map},
            {"mappingType": "discrete", "mappingColumn": "type", "mappingColumnType": "String",
             "visualProperty": "NODE_SHAPE",
             "map": [{"key": "药物", "value": "ELLIPSE"}, {"key": "成分", "value": "ELLIPSE"},
                     {"key": "靶点", "value": "ELLIPSE"}, {"key": "疾病", "value": "DIAMOND"}]},
            {"mappingType": "discrete", "mappingColumn": "type", "mappingColumnType": "String",
             "visualProperty": "NODE_SIZE",
             "map": [{"key": "药物", "value": "62"}, {"key": "成分", "value": "18"},
                     {"key": "靶点", "value": "22"}, {"key": "疾病", "value": "76"}]},
            {"mappingType": "discrete", "mappingColumn": "type", "mappingColumnType": "String",
             "visualProperty": "NODE_LABEL_FONT_SIZE",
             "map": [{"key": "药物", "value": "13"}, {"key": "成分", "value": "6"},
                     {"key": "靶点", "value": "7"}, {"key": "疾病", "value": "13"}]},
            {"mappingType": "discrete", "mappingColumn": "type", "mappingColumnType": "String",
             "visualProperty": "EDGE_TRANSPARENCY",
             "map": [{"key": "药物-成分", "value": "135"},
                     {"key": "成分-靶点", "value": "72"},
                     {"key": "靶点-疾病", "value": "24"}]},
            {"mappingType": "discrete", "mappingColumn": "type", "mappingColumnType": "String",
             "visualProperty": "EDGE_WIDTH",
             "map": [{"key": "药物-成分", "value": "0.75"},
                     {"key": "成分-靶点", "value": "0.48"},
                     {"key": "靶点-疾病", "value": "0.28"}]},
        ],
    }
    style_title = style_payload["title"]
    try:
        r = rest("POST", "/styles", style_payload)
        if isinstance(r, dict) and r.get("title"):
            style_title = r["title"]
            print("样式已创建:", style_title)
        else:
            print("样式创建返回:", r)
    except Exception as e:
        print("样式可能已存在, 直接应用(提示:", e, ")")
    try:
        rest("POST", "/commands/vizmap/apply",
             {"network": f"SUID:{suid}", "styles": style_title})
        print("样式已应用:", style_title)
    except Exception as e:
        print("样式应用失败(继续):", e)

    # 固定模板布局 + 适配视图（不再使用会破坏分区结构的 CoSE）。
    try:
        views = rest("GET", f"/networks/{suid}/views")
        view_id = views[0] if isinstance(views, list) and views else None
        if view_id:
            apply_positions(suid, view_id, nodes, template_positions(nodes))
            rest("POST", "/commands/view/fit%20content", {"view": f"SUID:{view_id}"})
            print(f"视图 {view_id} 已适配(fit content)")
    except Exception as e:
        print("模板布局或视图适配失败(继续):", e)

    # 从 Cytoscape 视图直接导出高分辨率 PNG 与矢量 SVG。
    try:
        views = rest("GET", f"/networks/{suid}/views")
        view_id = views[0] if isinstance(views, list) and views else None
        if not view_id:
            print("未找到视图")
        else:
            png_bytes = download(f"/networks/{suid}/views/{view_id}.png?w=5000", PNG_OUT)
            svg_bytes = download(f"/networks/{suid}/views/{view_id}.svg", SVG_OUT)
            print(f"已导出 {PNG_OUT} ({png_bytes} bytes)")
            print(f"已导出 {SVG_OUT} ({svg_bytes} bytes)")
    except Exception as e:
        print("导出失败:", e)

    # 保存 Cytoscape 会话，便于继续手工微调节点。
    try:
        abs_session = os.path.abspath(SESSION_OUT).replace("\\", "/")
        rest("POST", "/commands/session/save%20as", {"file": abs_session}, timeout=180)
        print(f"已保存 Cytoscape 会话: {SESSION_OUT}")
    except Exception as e:
        print("会话保存失败(图片不受影响):", e)

if __name__ == "__main__":
    main()
