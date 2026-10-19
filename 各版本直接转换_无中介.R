library(readxl)
library(dplyr)
library(writexl)
library(tibble)

# ── 跨平台文件定位（兼容 Windows / macOS / 不同用户目录）──────────────────────
# 获取脚本自身所在目录：Rscript 经 --file= 传入；RStudio 交互式回退工作目录
script_dir <- function() {
  fa <- commandArgs(trailingOnly = FALSE)
  hit <- fa[grepl("^--file=", fa)]
  if (length(hit)) {
    return(dirname(normalizePath(sub("^--file=", "", hit[1]))))
  }
  getwd()
}

# 读取命令行参数 --key=value
get_arg <- function(key, default = "") {
  args <- commandArgs(trailingOnly = TRUE)
  hit <- grep(paste0("^--", key, "="), args, value = TRUE)
  if (length(hit)) sub(paste0("^--", key, "="), "", hit[1]) else default
}

# 定位年代表：--xlsx= > 环境变量 GEO_XLSX > 脚本同目录候选 > 用户桌面 WJY2 回退
find_geology_xlsx <- function() {
  dir <- script_dir()
  candidates <- c(
    get_arg("xlsx"),
    Sys.getenv("GEO_XLSX", unset = ""),
    file.path(dir, "docs", "地质年代表.xlsx"),
    file.path(dir, "地质年代表.xlsx"),
    file.path(dir, "地质年代表（表头）.xlsx"),
    file.path(dir, "wjy.xlsx"),
    file.path(path.expand("~"), "Desktop", "WJY2", "地质年代表（表头）.xlsx"),
    file.path(path.expand("~"), "Desktop", "WJY2", "wjy.xlsx")
  )
  candidates <- candidates[nzchar(candidates)]
  for (p in candidates) {
    if (file.exists(p)) return(normalizePath(p))
  }
  cat("未找到年代表文件，已按以下顺序查找：\n")
  for (p in candidates) cat("  -", p, "\n")
  nearby <- list.files(dir, pattern = "\\.xlsx$")
  if (length(nearby)) cat("脚本目录内的 xlsx 文件：", paste(nearby, collapse = ", "), "\n")
  stop("请用 --xlsx=文件路径 指定年代表，或设置环境变量 GEO_XLSX")
}

# 在年代表中查找年龄对应的目标列值（支持精确匹配与线性插值）
# 通用查找：在 col_name 列中定位 target_age，返回同一行 return_col 列的值
find_age_in_column <- function(target_age, df, col_name, return_col) {
  col_vals <- suppressWarnings(as.numeric(df[[col_name]]))
  valid <- !is.na(col_vals)
  if (sum(valid) == 0) return(NULL)
  col_filt <- col_vals[valid]
  df_filt <- df[valid, ]
  # 精确匹配
  match_idx <- which(col_filt == target_age)[1]
  if (!is.na(match_idx)) {
    res <- suppressWarnings(as.numeric(as.character(df_filt[[return_col]][match_idx])))
    if (!is.na(res)) return(list(val=res, exact=TRUE))
  }
  # 插值
  lower_idx <- which(col_filt <= target_age)
  upper_idx <- which(col_filt >= target_age)
  if (length(lower_idx) == 0 || length(upper_idx) == 0) return(NULL)
  lower_idx <- lower_idx[which.max(col_filt[lower_idx])]
  upper_idx <- upper_idx[which.min(col_filt[upper_idx])]
  # 上下邻指向同一行 = 目标年龄恰为某控制点，此情形已被上面的 == 精确匹配覆盖，
  # 故这里直接返回 NULL（原先的 0.001 容差分支永远够不着，已删除）
  if (lower_idx == upper_idx) return(NULL)
  lower_x <- col_filt[lower_idx]
  upper_x <- col_filt[upper_idx]
  lower_y <- suppressWarnings(as.numeric(as.character(df_filt[[return_col]][lower_idx])))
  upper_y <- suppressWarnings(as.numeric(as.character(df_filt[[return_col]][upper_idx])))
  if (is.na(lower_y) || is.na(upper_y)) return(NULL)
  ratio <- (target_age - lower_x) / (upper_x - lower_x)
  interp <- lower_y + (upper_y - lower_y) * ratio
  list(val=interp, exact=FALSE)
}
# 处理单个年代：直接从 source_ver 转换到 target_ver（已去掉 IUGS 中介环节）
# 原为两步 source->IUGS 2023.9->target（两次插值，误差累积）；
# 现为一步 source->target（至多一次插值，精度更高，且不依赖任何中介列）。
convert_age <- function(age, df, source_ver, target_ver) {
  age_num <- suppressWarnings(as.numeric(age))
  if (is.na(age_num)) return(NULL)
  # 一步：source_ver -> target_ver
  step <- find_age_in_column(age_num, df, source_ver, target_ver)
  if (is.null(step)) {
    cat(sprintf("  无法在%s中定位%.4f并转换到%s\n", source_ver, age_num, target_ver))
    return(NULL)
  }
  final_age <- step$val
  cat(sprintf("  %.4f (%s) -> %.4f (%s) [%s]\n",
              age_num, source_ver, final_age, target_ver,
              if(step$exact) "精确" else "插值"))
  list(
    source = age_num,
    target = final_age,
    exact  = step$exact
  )
}
# 处理并返回结果行（地层信息按源列就近匹配，用于展示该年龄所属单元）
process <- function(age, df, source_ver, target_ver, show_cols) {
  res <- convert_age(age, df, source_ver, target_ver)
  if (is.null(res)) return(NULL)
  # 找显示行：用源版本列的值就近匹配，定位该年龄所属的地层单元
  src_vals <- suppressWarnings(as.numeric(as.character(df[[source_ver]])))
  nearest <- which.min(abs(src_vals - res$source))[1]
  result_row <- df[nearest, show_cols, drop=FALSE]
  result_row[[source_ver]] <- as.character(res$source)
  result_row[[target_ver]] <- as.character(round(res$target, 2))
  result_row[["转换方式"]] <- if(res$exact) "精确匹配" else "插值"
  return(result_row)
}
# 主函数
main <- function() {
  xlsx_path <- find_geology_xlsx()
  cat("年代表：", xlsx_path, "\n")
  # sheet 自适应：优先 Sheet1，不存在则取第一个工作表
  sheets <- readxl::excel_sheets(xlsx_path)
  use_sheet <- if ("Sheet1" %in% sheets) "Sheet1" else sheets[1]
  df <- read_excel(xlsx_path, sheet = use_sheet, col_types = "text")
  cat(sprintf("数据: %d行 x %d列\n", nrow(df), ncol(df)))
  # 显示可用版本（排除误差列和空列）
  all_cols <- names(df)
  version_candidates <- all_cols[!grepl("误差|error|Unnamed", all_cols, ignore.case=TRUE)]
  version_candidates <- version_candidates[version_candidates != ""]
  cat("\n可用版本:\n")
  for (i in seq_along(version_candidates)) {
    cat(sprintf("  %d. %s\n", i, version_candidates[i]))
  }
  # 选择源版本
  source_input <- trimws(readline("\n源版本（输入列名或编号）: "))
  source_ver <- if (suppressWarnings(!is.na(as.numeric(source_input)))) {
    version_candidates[as.numeric(source_input)]
  } else {
    source_input
  }
  if (!(source_ver %in% names(df))) {
    cat("源版本不存在\n")
    return(NULL)
  }
  cat(sprintf("源版本: %s\n", source_ver))
  # 选择目标版本
  target_input <- trimws(readline("目标版本（输入列名或编号）: "))
  target_ver <- if (suppressWarnings(!is.na(as.numeric(target_input)))) {
    version_candidates[as.numeric(target_input)]
  } else {
    target_input
  }
  if (!(target_ver %in% names(df))) {
    cat("目标版本不存在\n")
    return(NULL)
  }
  if (target_ver == source_ver) {
    cat("源版本和目标版本相同，无需转换\n")
    return(NULL)
  }
  cat(sprintf("目标版本: %s\n", target_ver))
  # ── 模式选择：处理表格 or 直接输入单个/多个年龄 ──
  cat("\n请选择更新方式:\n")
  cat("  1. 处理表格（读取文件 → 批量更新 → 写回原文件）\n")
  cat("  2. 直接输入年龄（即时转换并打印结果，不读写任何文件）\n")
  mode_input <- trimws(readline("输入 1 或 2 [直接回车默认 1]: "))
  if (mode_input == "") mode_input <- "1"
  # 显示列：地层信息 + 源列 + 目标列（两种模式共用；已移除 IUGS 中介列）
  show_cols <- c('宇','界','系','统','阶',
                 'Eonothem','Erathem','System','Series','Stage',
                 source_ver, target_ver)
  show_cols <- show_cols[show_cols %in% names(df)]
  # 就近匹配源列，定位某年龄所属的地层单元名（两种模式共用）
  stage_of <- function(a) {
    src_vals <- suppressWarnings(as.numeric(as.character(df[[source_ver]])))
    nearest <- which.min(abs(src_vals - a))[1]
    for (cn in c('阶', 'Stage')) {
      if (cn %in% names(df)) return(as.character(df[[cn]][nearest]))
    }
    return("")
  }
  if (mode_input == "2") {
    # ===================== 模式 2：直接输入年龄 =====================
    cat("\n直接输入年龄，多个用空格或逗号分隔（例：251.4 250.7）\n")
    raw <- trimws(readline("年龄: "))
    tokens <- unlist(strsplit(raw, "[,，;；、[:space:]]+"))
    tokens <- tokens[tokens != ""]
    ages <- suppressWarnings(as.numeric(tokens))
    ages <- ages[!is.na(ages)]
    if (length(ages) == 0) {
      cat("未解析到有效年龄\n")
      return(NULL)
    }
    cat(sprintf("\n将转换 %d 个年龄: %s -> %s（直接转换，无中介）\n",
                length(ages), source_ver, target_ver))
    cat(strrep("-", 70), "\n")
    ok <- 0
    for (a in ages) {
      res <- convert_age(a, df, source_ver, target_ver)
      if (is.null(res)) {
        cat(sprintf("  %-14s 无法转换（超出 %s 控制点范围或无匹配行）\n",
                    format(a, nsmall = 2), source_ver))
        next
      }
      ok <- ok + 1
      cat(sprintf("  %-14s -> %-14s [%s]  所属单元: %s\n",
                  format(res$source, nsmall = 2),
                  format(res$target, nsmall = 2),
                  if (res$exact) "精确匹配" else "插值",
                  stage_of(a)))
    }
    cat(strrep("-", 70), "\n")
    cat(sprintf("成功转换 %d/%d 个年龄（结果仅打印，未写入文件）\n", ok, length(ages)))
    return(NULL)
  }
  # ===================== 模式 1：处理表格（原有逻辑） =====================
  # 读取待处理文件
  file_path <- trimws(readline("\n待处理文件: "))
  if (!file.exists(file_path)) {
    cat("文件不存在\n")
    return(NULL)
  }
  src_df <- read_excel(file_path, col_types="text")
  cat("文件列:", paste(names(src_df), collapse=" | "), "\n")
  age_col <- trimws(readline("年代数据列名: "))
  if (!(age_col %in% names(src_df))) {
    cat("列不存在\n")
    return(NULL)
  }
  ages <- suppressWarnings(as.numeric(src_df[[age_col]]))
  valid_idx <- which(!is.na(ages))
  ages <- ages[valid_idx]
  if (length(ages) == 0) {
    cat("没有有效年代数据\n")
    return(NULL)
  }
  cat(sprintf("\n将转换 %d 个年代: %s -> %s（直接转换，无中介）\n",
              length(ages), source_ver, target_ver))
  # 处理
  results <- list()
  result_map <- list()
  for (a in ages) {
    r <- process(a, df, source_ver, target_ver, show_cols)
    if (!is.null(r)) {
      results[[length(results)+1]] <- r
      result_map[[as.character(a)]] <- list(
        target = r[[target_ver]][1],
        method = r[["转换方式"]][1]
      )
    }
  }
  if (length(results) == 0) {
    cat("无成功转换的数据\n")
    return(NULL)
  }
  cat(sprintf("\n成功转换 %d/%d 个年代\n", length(results), length(ages)))
  # 写回结果
  new_col <- paste0(target_ver, " (来自", source_ver, ")")
  counter <- 1
  while (new_col %in% names(src_df)) {
    new_col <- paste0(target_ver, " (来自", source_ver, ")_", counter)
    counter <- counter + 1
  }
  # 初始化新列（已移除 _IUGS中介 列）
  src_df[[new_col]] <- NA
  src_df[[paste0(new_col, "_方式")]] <- NA
  # 填充数据
  for (i in seq_along(src_df[[age_col]])) {
    orig_age <- suppressWarnings(as.numeric(as.character(src_df[[age_col]][i])))
    if (!is.na(orig_age)) {
      key <- as.character(orig_age)
      if (!is.null(result_map[[key]])) {
        src_df[[new_col]][i] <- result_map[[key]]$target
        src_df[[paste0(new_col, "_方式")]][i] <- result_map[[key]]$method
      }
    }
  }
  write_xlsx(src_df, file_path)
  cat(sprintf("\n已写入 '%s':\n", basename(file_path)))
  cat(sprintf("  - %s: 转换后的目标年代（直接由 %s 转换）\n", new_col, source_ver))
  cat(sprintf("  - %s_方式: 精确匹配或插值\n", new_col))
}
main()
