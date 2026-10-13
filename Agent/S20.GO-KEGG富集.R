# -*- coding: utf-8 -*-
# S20.GO-KEGG富集.R
# =================
# 对核心基因(默认 S19.核心基因列表.txt) 做 GO(BP/CC/MF) 与 KEGG 富集分析
# 依赖: R + clusterProfiler, org.Hs.eg.db, enrichplot, ggplot2, openxlsx
#
# 用法:
#   Rscript S20.GO-KEGG富集.R [--genes 基因列表.txt] [--prefix S20] [--pvalue 0.05]
#                            [--qvalue 0.05] [--bg 背景基因.txt]
# 背景: 默认使用 org.Hs.eg.db 全部基因; 可用 --bg 指定更保守背景(如 PPI 网络基因)
# 输出: {prefix}.GO-KEGG富集.xlsx + 气泡图/柱状图 PNG
suppressMessages({
  library(clusterProfiler); library(org.Hs.eg.db); library(enrichplot)
  library(ggplot2); library(openxlsx)
})

# KEGG 在线注释在网络波动时可能超过 R 默认 60 秒。
options(timeout = 300)

args <- commandArgs(trailingOnly = TRUE)
get_arg <- function(key, default) {
  i <- match(key, args)
  if (is.na(i)) default else args[i + 1]
}
gene_file <- get_arg("--genes", "S19.核心基因列表.txt")
prefix    <- get_arg("--prefix", "S20")
pval      <- as.numeric(get_arg("--pvalue", "0.05"))
qval      <- as.numeric(get_arg("--qvalue", "0.05"))
bg_file   <- get_arg("--bg", "")

genes <- readLines(gene_file, warn = FALSE)
genes <- trimws(genes[nzchar(genes)])
cat("输入基因数:", length(genes), "\n")

mapping <- bitr(genes, fromType = "SYMBOL", toType = "ENTREZID", OrgDb = org.Hs.eg.db)
entrez <- unique(mapping$ENTREZID)
cat("映射到 EntrezID:", length(entrez), "\n")
if (length(entrez) == 0) stop("无有效 EntrezID")

# 背景基因集
universe <- NULL
if (nzchar(bg_file)) {
  bg <- readLines(bg_file, warn = FALSE)
  bg <- trimws(bg[nzchar(bg)])
  bg_map <- bitr(bg, fromType = "SYMBOL", toType = "ENTREZID", OrgDb = org.Hs.eg.db)
  universe <- unique(bg_map$ENTREZID)
  cat("使用自定义背景基因:", length(universe), "\n")
}

run_enrich <- function(expr) tryCatch(expr, error = function(e) {
  cat("  [警告] 富集失败:", conditionMessage(e), "\n"); NULL })

cat("运行 GO 富集 (BP/CC/MF) ...\n")
go_all <- run_enrich(enrichGO(gene = entrez, universe = universe, OrgDb = org.Hs.eg.db,
                              ont = "ALL", pAdjustMethod = "BH",
                              pvalueCutoff = pval, qvalueCutoff = qval, readable = TRUE))
go_bp <- run_enrich(enrichGO(gene = entrez, universe = universe, OrgDb = org.Hs.eg.db,
                             ont = "BP", pAdjustMethod = "BH",
                             pvalueCutoff = pval, qvalueCutoff = qval, readable = TRUE))
go_cc <- run_enrich(enrichGO(gene = entrez, universe = universe, OrgDb = org.Hs.eg.db,
                             ont = "CC", pAdjustMethod = "BH",
                             pvalueCutoff = pval, qvalueCutoff = qval, readable = TRUE))
go_mf <- run_enrich(enrichGO(gene = entrez, universe = universe, OrgDb = org.Hs.eg.db,
                             ont = "MF", pAdjustMethod = "BH",
                             pvalueCutoff = pval, qvalueCutoff = qval, readable = TRUE))
cat("运行 KEGG 富集 (hsa, 需联网) ...\n")
kegg <- run_enrich(enrichKEGG(gene = entrez, universe = universe, organism = "hsa",
                              pvalueCutoff = pval, qvalueCutoff = qval))
if (!is.null(kegg) && nrow(as.data.frame(kegg)) > 0)
  kegg <- setReadable(kegg, OrgDb = org.Hs.eg.db, keyType = "ENTREZID")

# ---- 输出 xlsx ----
out_xlsx <- paste0(prefix, ".GO-KEGG富集.xlsx")
wb <- createWorkbook()
addWorksheet(wb, "输入基因")
writeData(wb, "输入基因", data.frame(GeneSymbol = genes,
                                     EntrezID = mapping$ENTREZID[match(genes, mapping$SYMBOL)]))
n_sum <- data.frame(项目 = character(), 值 = integer(), stringsAsFactors = FALSE)
for (nm in c("GO_BP", "GO_CC", "GO_MF", "KEGG")) {
  obj <- switch(nm, GO_BP = go_bp, GO_CC = go_cc, GO_MF = go_mf, KEGG = kegg)
  addWorksheet(wb, nm)
  if (!is.null(obj) && nrow(as.data.frame(obj)) > 0) {
    d <- as.data.frame(obj)
    writeData(wb, nm, d)
    n_sum <- rbind(n_sum, data.frame(项目 = nm, 值 = nrow(d)))
    cat(nm, "显著条目:", nrow(d), "\n")
  } else {
    writeData(wb, nm, data.frame(Note = "无显著富集条目"))
    n_sum <- rbind(n_sum, data.frame(项目 = nm, 值 = 0))
    cat(nm, "无显著条目\n")
  }
}
addWorksheet(wb, "富集统计")
writeData(wb, "富集统计", rbind(
  data.frame(项目 = c("输入基因数", "映射Entrez数", "pvalueCutoff", "qvalueCutoff",
                     "背景基因数", "物种", "GO注释库", "KEGG注释库"),
             值 = c(length(genes), length(entrez), pval, qval,
                       if (is.null(universe)) "默认(org.Hs.eg.db全部)" else length(universe),
                       "Homo sapiens (hsa)", "org.Hs.eg.db", "KEGG在线(hsa)")),
  n_sum))
saveWorkbook(wb, out_xlsx, overwrite = TRUE)
cat("已写出", out_xlsx, "\n")

# ---- 绘图 ----
plot_dot <- function(obj, file, title) {
  if (is.null(obj) || nrow(as.data.frame(obj)) == 0) { cat("跳过绘图(无条目):", file, "\n"); return() }
  p <- tryCatch(dotplot(obj, showCategory = 15) + ggtitle(title), error = function(e) NULL)
  if (!is.null(p)) { ggsave(file, p, width = 10, height = 9, dpi = 300); cat("已写出", file, "\n") }
}
plot_bar <- function(obj, file, title) {
  if (is.null(obj) || nrow(as.data.frame(obj)) == 0) { cat("跳过绘图(无条目):", file, "\n"); return() }
  p <- tryCatch(barplot(obj, showCategory = 15) + ggtitle(title), error = function(e) NULL)
  if (!is.null(p)) { ggsave(file, p, width = 10, height = 9, dpi = 300); cat("已写出", file, "\n") }
}

# GO 合并图 (BP/CC/MF facet)
if (!is.null(go_all) && nrow(as.data.frame(go_all)) > 0) {
  p1 <- dotplot(go_all, showCategory = 10) + facet_grid(ONTOLOGY ~ ., scale = "free") +
        ggtitle("GO Enrichment (Core genes)")
  ggsave(paste0(prefix, ".GO气泡图.png"), p1, width = 11, height = 12, dpi = 300)
  cat("已写出", paste0(prefix, ".GO气泡图.png"), "\n")
  p2 <- barplot(go_all, showCategory = 10) + facet_grid(ONTOLOGY ~ ., scale = "free") +
        ggtitle("GO Enrichment (Core genes)")
  ggsave(paste0(prefix, ".GO柱状图.png"), p2, width = 11, height = 12, dpi = 300)
  cat("已写出", paste0(prefix, ".GO柱状图.png"), "\n")
}
plot_dot(go_bp, paste0(prefix, ".GO-BP气泡图.png"), "GO Biological Process")
plot_dot(go_cc, paste0(prefix, ".GO-CC气泡图.png"), "GO Cellular Component")
plot_dot(go_mf, paste0(prefix, ".GO-MF气泡图.png"), "GO Molecular Function")
plot_dot(kegg, paste0(prefix, ".KEGG气泡图.png"), "KEGG Pathway")
plot_bar(kegg, paste0(prefix, ".KEGG柱状图.png"), "KEGG Pathway")

cat("=== S20 GO-KEGG 富集完成 ===\n")
