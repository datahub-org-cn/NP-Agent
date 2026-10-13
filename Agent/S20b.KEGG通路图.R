# -*- coding: utf-8 -*-
# S20b.KEGG通路图.R
# 对 S20 富集得到的显著 KEGG 通路, 用 pathview 绘制通路图并高亮核心基因
suppressMessages({
  library(clusterProfiler); library(org.Hs.eg.db); library(pathview)
})

# KEGG KGML/底图下载在网络波动时可能超过 R 默认 60 秒。
options(timeout = 300)

args <- commandArgs(trailingOnly = TRUE)
get_arg <- function(key, default) { i <- match(key, args); if (is.na(i)) default else args[i+1] }
gene_file <- get_arg("--genes", "S19.核心基因列表.txt")
score_xlsx <- get_arg("--score", "S19.核心基因.xlsx")
kegg_xlsx  <- get_arg("--kegg", "S20.GO-KEGG富集.xlsx")
top_n      <- as.integer(get_arg("--top", "10"))
suffix     <- get_arg("--suffix", "core")

genes <- readLines(gene_file, warn = FALSE); genes <- trimws(genes[nzchar(genes)])
map <- bitr(genes, fromType = "SYMBOL", toType = "ENTREZID", OrgDb = org.Hs.eg.db)
cat("核心基因:", length(genes), "映射 Entrez:", nrow(map), "\n")

# 读取 CS 评分(可选) -> 基因着色深浅
cs <- rep(1, nrow(map)); names(cs) <- map$ENTREZID
if (file.exists(score_xlsx)) {
  suppressMessages({ library(readxl) })
  d <- tryCatch(as.data.frame(readxl::read_excel(score_xlsx, sheet = "核心基因评分")), error = function(e) NULL)
  if (!is.null(d) && "基因" %in% names(d) && "CS" %in% names(d)) {
    m <- merge(data.frame(SYMBOL = map$SYMBOL, ENTREZID = map$ENTREZID), d, by.x = "SYMBOL", by.y = "基因", all.x = TRUE)
    cs <- setNames(m$CS, m$ENTREZID)
    cs[is.na(cs)] <- 1
    cat("已使用 CS 评分着色\n")
  }
}

# 读取 KEGG 显著通路
suppressMessages({ library(readxl) })
kg <- tryCatch(as.data.frame(readxl::read_excel(kegg_xlsx, sheet = "KEGG")), error = function(e) NULL)
if (is.null(kg) || !("ID" %in% names(kg))) { cat("无法读取 KEGG 富集表\n"); quit(status = 1) }
kg$p.adjust <- as.numeric(kg$p.adjust)
kg <- kg[!is.na(kg$p.adjust), ]
kg <- kg[order(kg$p.adjust), ]
kg <- head(kg, top_n)
cat("绘制前", nrow(kg), "条通路:\n")
print(data.frame(ID = kg$ID, Description = substr(kg$Description, 1, 60), p.adjust = signif(kg$p.adjust, 3)))

made <- c()
for (i in seq_len(nrow(kg))) {
  pid <- kg$ID[i]
  desc <- kg$Description[i]
  cat("  [", i, "/", nrow(kg), "] ", pid, " ", substr(desc, 1, 40), " ... ", sep = "")
  ok <- tryCatch({
    pv <- pathview(gene.data = cs, pathway.id = pid, species = "hsa",
                   out.suffix = suffix, kegg.native = TRUE, same.layer = TRUE)
    file.exists(paste0(pid, ".", suffix, ".png"))
  }, error = function(e) { cat("失败:", conditionMessage(e), "\n"); FALSE })
  if (ok) { made <- c(made, paste0(pid, ".", suffix, ".pathview.png")); cat("OK\n") } else cat("无输出\n")
}
cat("生成通路图:\n"); print(made)
cat("=== S20b KEGG 通路图完成 ===\n")
