# NP-Agent

网络药理学分析流程，包含 Python / R 分步脚本，以及逍遥丸干预抑郁障碍案例的输入、分析结果和图表。

## 仓库内容

- `Agent/`：保留原始文件名和相对路径的 S1–S21 流程文件。
- [研究说明](Agent/研究说明.md)：研究边界、参数、结果统计及文件索引。
- `requirements.txt`：从 Python 源码导入项整理的第三方依赖，未锁定版本。

| 阶段 | 内容 |
| --- | --- |
| S1–S5 | 复方输入、HERB / TCMSP 成分获取、活性筛选、假阳性排除和去重 |
| S6–S12 | 药物与成分靶点、Gene Symbol 补全、预测靶点和汇总 |
| S13–S16 | 疾病靶点、GeneCards 合并、人源校验和药物–疾病交集 |
| S17–S19 | Cytoscape 调控网络、STRING PPI 和核心基因筛选 |
| S20 / S20b | GO / KEGG 富集与通路图绘制 |
| S21 | 核心基因与成分对应 |

## 环境准备

安装 Python 3 和依赖：

```sh
python -m pip install -r requirements.txt
```

富集分析需要 R。在 R 中安装所需包：

```r
install.packages(c("BiocManager", "ggplot2", "openxlsx", "readxl"))
BiocManager::install(c("clusterProfiler", "org.Hs.eg.db", "enrichplot", "pathview"))
```

S17 绘图需要运行 Cytoscape，并使 cyREST 在 `http://127.0.0.1:1234` 可访问。其他涉及 Cytoscape 的步骤请参照脚本头部说明。

## 运行方式

先进入 `Agent` 目录，按步骤编号执行相应脚本；同一步骤的 HERB 与 TCMSP 脚本分别处理各自的数据来源。输入表包括 `S1复方表.xlsx`、`S13疾病.xlsx` 和 `S14.GeneCards.csv`。

```sh
cd Agent
python "S16.绘制Venn图.py"
Rscript "S20.GO-KEGG富集.R"
Rscript "S20b.KEGG通路图.R"
python "S21.核心基因-成分对应.py"
```

以上展示单步调用方式，并非从头执行的完整命令列表。各脚本的输入、输出和可选参数见脚本头部。仓库保留已有结果，重新运行可能覆盖同名输出。联网步骤依赖上游数据库的可用性与接口，重新查询所得结果可能随数据库更新发生变化。

## 收录范围与验证

本次上传收录 S1–S21 的脚本、配套表格、结果图、Cytoscape 会话、布局说明及研究说明。缓存目录 `cashe/`、Python 缓存、论文 HTML、测试报告、参考模板、旧失眠检索文件和可由 S20b 生成的 KEGG 底图 / XML 未收录。

上传前检查了 Python 脚本语法，并逐文件核对复制内容；未重新执行联网采集、Cytoscape 绘图或 R 富集分析。案例结果说明见原始研究说明。
