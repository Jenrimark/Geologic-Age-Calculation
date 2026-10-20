#!/usr/bin/env Rscript
# =============================================================================
#  年代表年龄转换 · 一体化 R 服务（网页 + HTTP/JSON API + Excel 批量上传）
#
#  端点：
#    GET  /                演示网页（R 同源托管，无跨域问题）
#    GET  /health          健康检查
#    GET  /versions        可用版本列表（序号可直接作 source/target）
#    GET/POST /convert     手动输入年龄转换（ages 数组或分隔字符串）
#    POST /inspect-file    上传 Excel -> 返回工作表与列名，供前端选择
#    POST /convert-file    上传 Excel -> 批量转换 -> 返回结果文件(base64 下载)
#    GET  /shutdown        停止服务
#
#  相比早期版本的改进：
#    1. 支持网页上传 Excel 批量转换并下载结果，结果以下载返回，绝不覆盖原文件
#    2. 修复 get_arg 缺省参数缺失导致不带 --xlsx= 就崩溃的问题
#    3. 修复 wb$getSheetNames() 在 openxlsx 4.2.x 不存在导致的多工作表误写第一个表
#    4. 保存前清理悬空 drawing/vml 引用，避免含批注源文件产出后 Excel 报「文件已损坏」
#    5. 保存前修正 workbookView 空属性，保证 openpyxl 等严格解析器可打开
#    6. 端口被占用时自动顺延，出错时窗口停住显示原因（双击启动不再闪退）
#
#  启动：
#    Rscript 年代表转换_文件上传服务.R --xlsx="年代表路径" [--port=8788] [--host=127.0.0.1]
#
#  依赖：httpuv jsonlite readxl openxlsx writexl（zip 非必需）
# =============================================================================

suppressWarnings(suppressMessages({
  library(httpuv)
  library(jsonlite)
  library(readxl)
  library(openxlsx)
  library(writexl)
}))

# 双击启动时窗口一闪而过就什么也看不到，所以出错先打印原因、再等用户按键。
options(error = function() {
  cat("\n")
  cat(strrep("-", 66), "\n")
  cat("启动失败。上面的错误信息即原因，请截图保存。\n")
  cat("常见原因：R 缺少 httpuv/jsonlite/readxl/openxlsx/writexl 包，\n")
  cat("          或找不到年代表文件、找不到「各版本直接转换_无中介.R」。\n")
  cat(strrep("-", 66), "\n")
  cat("按回车键关闭本窗口…")
  invisible(tryCatch(readLines(stdin(), n = 1), error = function(e) NULL))
})

# ── 命令行参数 ────────────────────────────────────────────────────────────────
args <- commandArgs(trailingOnly = TRUE)
# 注意：default 必须有缺省值。归档版写成 get_arg("xlsx") 缺第二参，
# 不传 --xlsx= 时因 R 惰性求值触发 "argument default is missing" 而崩溃。
get_arg <- function(key, default = "") {
  hit <- grep(paste0("^--", key, "="), args, value = TRUE)
  if (length(hit)) sub(paste0("^--", key, "="), "", hit[1]) else default
}

SCRIPT_DIR <- (function() {
  fa <- commandArgs(trailingOnly = FALSE)
  hit <- fa[grepl("^--file=", fa)]
  if (length(hit)) dirname(normalizePath(sub("^--file=", "", hit[1]))) else getwd()
})()

WEB_DIR <- file.path(SCRIPT_DIR, "web")
if (!dir.exists(WEB_DIR)) WEB_DIR <- SCRIPT_DIR

# 用户主目录：Windows 下 path.expand("~") 可能被 R_USER/HOME 带偏到「文档」，
# 故优先 USERPROFILE，再退回 HOMEDRIVE+HOMEPATH，最后才用 path.expand("~")。
USER_HOME <- (function() {
  cand <- c(
    Sys.getenv("USERPROFILE", unset = ""),
    (function() {
      d <- Sys.getenv("HOMEDRIVE", unset = ""); p <- Sys.getenv("HOMEPATH", unset = "")
      if (nzchar(d) && nzchar(p)) paste0(d, p) else ""
    })(),
    tryCatch(normalizePath(path.expand("~"), winslash = "/"), error = function(e) "")
  )
  cand <- cand[nzchar(cand)]
  hit <- cand[dir.exists(cand)]
  if (length(hit)) return(hit[1])
  getwd()
})()

# 核心转换脚本：--core= > GEO_CORE > 同目录 > 桌面 WJY2（日常维护版）> 桌面归档
# 优先 WJY2：那是用户日常编辑的位置，更新后网页自动生效；归档仅作回退，保持不动。
resolve_core <- function() {
  cand <- c(
    get_arg("core"),
    Sys.getenv("GEO_CORE", unset = ""),
    file.path(SCRIPT_DIR, "各版本直接转换_无中介.R"),
    file.path(USER_HOME, "Desktop", "WJY2", "各版本直接转换_无中介.R"),
    file.path(USER_HOME, "Desktop", "归档", "各版本直接转换_无中介.R")
  )
  cand <- cand[nzchar(cand)]
  hit <- cand[file.exists(cand)]
  if (length(hit)) return(normalizePath(hit[1]))
  cat("未找到转换核心脚本，已按以下顺序查找：\n")
  for (p in cand) cat("  -", p, "\n")
  stop("请用 --core=路径 指定「各版本直接转换_无中介.R」")
}
USER_R <- resolve_core()

# 年代表：--xlsx= > GEO_XLSX > 同目录候选 > 桌面 WJY2（表头版，日常维护）> 归档 docs
resolve_xlsx <- function() {
  cand <- c(
    get_arg("xlsx"),
    Sys.getenv("GEO_XLSX", unset = ""),
    file.path(SCRIPT_DIR, "docs", "地质年代表.xlsx"),
    file.path(SCRIPT_DIR, "地质年代表.xlsx"),
    file.path(SCRIPT_DIR, "地质年代表（表头）.xlsx"),
    file.path(USER_HOME, "Desktop", "WJY2", "地质年代表（表头）.xlsx"),
    file.path(USER_HOME, "Desktop", "WJY2", "wjy.xlsx"),
    file.path(USER_HOME, "Desktop", "归档", "docs", "地质年代表.xlsx")
  )
  cand <- cand[nzchar(cand)]
  hit <- cand[file.exists(cand)]
  if (length(hit)) return(normalizePath(hit[1]))
  cat("未找到年代表，已按以下顺序查找：\n")
  for (p in cand) cat("  -", p, "\n")
  stop("请用 --xlsx=路径 指定年代表，或设置环境变量 GEO_XLSX")
}
XLSX <- resolve_xlsx()

PORT <- suppressWarnings(as.integer(get_arg("port", "8787")))
if (is.na(PORT) || PORT < 1 || PORT > 65535) PORT <- 8787L
HOST <- get_arg("host", "127.0.0.1")
if (!nzchar(HOST)) HOST <- "127.0.0.1"

MAX_BYTES <- suppressWarnings(as.integer(get_arg("maxmb", "40"))) * 1024 * 1024
if (is.na(MAX_BYTES)) MAX_BYTES <- 40 * 1024 * 1024

# ── 载入转换逻辑（剥离 library() 与 main() 调用，避免触发交互） ─────────────────
load_core <- function(path) {
  if (!file.exists(path)) stop("找不到核心脚本：", path)
  lines <- readLines(path, encoding = "UTF-8", warn = FALSE)
  keep <- !grepl("^\\s*library\\(", lines) & !grepl("^\\s*main\\(\\)\\s*$", lines)
  env <- new.env(parent = globalenv())
  eval(parse(text = paste(lines[keep], collapse = "\n")), envir = env)
  for (fn in c("find_age_in_column", "convert_age")) {
    if (!is.function(get0(fn, envir = env))) {
      stop("核心脚本中未找到函数 ", fn, "，无法提供转换服务")
    }
  }
  env
}
cat("载入转换逻辑：", USER_R, "\n")
CORE <- load_core(USER_R)
convert_age <- CORE$convert_age

# ── 读取年代表（启动时一次，常驻内存，只读） ──────────────────────────────────
cat("读取年代表：", XLSX, "\n", sep = "")
sheets0 <- readxl::excel_sheets(XLSX)
use_sheet0 <- if ("Sheet1" %in% sheets0) "Sheet1" else sheets0[1]
DF <- read_excel(XLSX, sheet = use_sheet0, col_types = "text")
cat(sprintf("  %d 行 x %d 列（工作表：%s）\n", nrow(DF), ncol(DF), use_sheet0))

HIER_COLS <- c("Eonothem", "Erathem", "System", "Series", "Stage",
               "宇", "界", "系", "统", "阶")
ALL_COLS <- names(DF)
CANDIDATES <- ALL_COLS[!grepl("误差|error|Unnamed", ALL_COLS, ignore.case = TRUE)]
CANDIDATES <- CANDIDATES[nzchar(CANDIDATES)]
CANDIDATES <- CANDIDATES[!(CANDIDATES %in% HIER_COLS)]
cat(sprintf("  可用版本 %d 个\n", length(CANDIDATES)))

# ── 工具函数 ──────────────────────────────────────────────────────────────────
`%||%` <- function(a, b) if (is.null(a)) b else a

url_decode <- function(s) {
  if (is.null(s) || length(s) == 0 || is.na(s)) return("")
  s <- gsub("+", " ", s, fixed = TRUE)
  if (!grepl("%", s, fixed = TRUE)) return(s)
  parts <- strsplit(s, "%", fixed = TRUE)[[1]]
  bytes <- charToRaw(parts[1])
  if (length(parts) > 1) {
    for (p in parts[-1]) {
      if (nchar(p) >= 2 && grepl("^[0-9A-Fa-f]{2}", p)) {
        v <- suppressWarnings(strtoi(substr(p, 1, 2), 16L))
        if (!is.na(v) && v >= 0 && v <= 255) bytes <- c(bytes, as.raw(v))
        bytes <- c(bytes, charToRaw(substr(p, 3, nchar(p))))
      } else {
        bytes <- c(bytes, charToRaw(paste0("%", p)))
      }
    }
  }
  res <- tryCatch(rawToChar(bytes), error = function(e) NULL)
  if (is.null(res)) return(s)
  Encoding(res) <- "UTF-8"
  res
}

parse_qs <- function(qs) {
  out <- list()
  if (is.null(qs) || !nzchar(qs)) return(out)
  qs <- sub("^\\?", "", qs)
  for (pair in strsplit(qs, "&", fixed = TRUE)[[1]]) {
    if (!nzchar(pair)) next
    eq <- regexpr("=", pair, fixed = TRUE)
    if (eq < 1) {
      out[[url_decode(pair)]] <- ""
    } else {
      out[[url_decode(substr(pair, 1, eq - 1))]] <-
        url_decode(substr(pair, eq + 1, nchar(pair)))
    }
  }
  out
}

json_resp <- function(obj, status = 200L) {
  list(
    status = status,
    headers = list(
      "Content-Type" = "application/json; charset=utf-8",
      "Access-Control-Allow-Origin" = "*",
      "Cache-Control" = "no-store"
    ),
    body = toJSON(obj, auto_unbox = TRUE, null = "null", na = "null")
  )
}

resolve_version <- function(v) {
  if (is.null(v) || is.na(v)) return(NULL)
  v <- trimws(as.character(v))
  if (!nzchar(v)) return(NULL)
  num <- suppressWarnings(as.numeric(v))
  if (!is.na(num) && num == floor(num) && num >= 1 && num <= length(CANDIDATES)) {
    return(CANDIDATES[num])
  }
  if (v %in% ALL_COLS) return(v)
  NULL
}

stage_of <- function(a, src_col) {
  vals <- suppressWarnings(as.numeric(as.character(DF[[src_col]])))
  if (all(is.na(vals))) return("")
  idx <- which.min(abs(vals - a))[1]
  for (cn in c("阶", "Stage")) {
    if (cn %in% names(DF)) {
      v <- DF[[cn]][idx]
      if (!is.na(v) && nzchar(trimws(as.character(v)))) return(as.character(v))
    }
  }
  ""
}

parse_ages <- function(x) {
  if (is.null(x)) return(numeric(0))
  if (is.list(x)) x <- unlist(x, use.names = FALSE)
  if (is.numeric(x)) {
    v <- x
  } else {
    s <- paste(as.character(x), collapse = " ")
    toks <- unlist(strsplit(s, "[,，;；、\\s]+"))
    toks <- toks[nzchar(trimws(toks))]
    v <- suppressWarnings(as.numeric(toks))
  }
  v[!is.na(v)]
}

# base64 -> 临时文件
b64_to_tmp <- function(b64, ext = ".xlsx") {
  b64 <- as.character(b64)[1]
  b64 <- sub("^data:[^,]*,", "", b64)          # 剥掉 data:...;base64, 前缀
  b64 <- gsub("[[:space:]]", "", b64)
  if (!nzchar(b64)) stop("空内容")
  raw <- base64_dec(b64)
  p <- tempfile(fileext = ext)
  con <- file(p, "wb"); on.exit(close(con), add = TRUE)
  writeBin(raw, con)
  p
}

tmp_to_b64 <- function(p) {
  raw <- readBin(p, what = "raw", n = file.info(p)$size)
  base64_enc(raw)
}

# 修正 workbookView 上的空属性（如 xWindow=""）。
# OOXML 规定这些属性必须是整数，空串不合规：Excel/WPS 通常容忍，
# 但严格按 schema 解析的库（如 Python openpyxl）会直接报错打不开。
# 空串来源：用户上传的文件本身缺该属性，openxlsx 读入后按空值写回。
# 必须在 saveWorkbook 之前修内存里的对象——事后重打包 zip 会破坏内部目录结构。
fix_bookviews <- function(wb) {
  tryCatch({
    bv <- wb$workbook$bookViews
    if (is.null(bv) || length(bv) == 0 || !nzchar(bv[1])) return(FALSE)
    fixed <- gsub(
      '\\b(xWindow|yWindow|windowWidth|windowHeight|firstSheet|activeTab|tabRatio)="\\s*"',
      '\\1="0"', bv[1], perl = TRUE)
    if (identical(fixed, bv[1])) return(FALSE)          # 无需修正
    wb$workbook$bookViews <- fixed
    TRUE
  }, error = function(e) {
    cat("bookViews 修正跳过：", conditionMessage(e), "\n"); FALSE
  })
}

# 清理 openxlsx 往返产生的悬空引用。
# openxlsx 4.2.x 的 loadWorkbook() 不解析批注部件（wb$comments 载入后是空 list），
# 但 worksheets_rels 里仍保留指向 drawing1.xml / vmlDrawing1.vml 的关系，
# 而这两个部件不会被打进包 —— 形成悬空引用。
# 后果：Excel 打开时提示「文件已损坏，是否修复」；严格按 schema 解析的库直接报错。
# 含批注的表格（如带填写说明、批注标注的表）必然触发，必须在保存前清理。
# 做法：删掉 drawing / vmlDrawing 关系，并清空 worksheet 上引用这些 rId 的标签。
clean_dangling <- function(wb) {
  tryCatch({
    n_rel <- 0L; n_tag <- 0L
    for (i in seq_along(wb$worksheets_rels)) {
      rel <- wb$worksheets_rels[[i]]
      if (!length(rel)) next
      keep <- !grepl('relationships/(drawing|vmlDrawing)"', rel)
      if (any(!keep)) { wb$worksheets_rels[[i]] <- rel[keep]; n_rel <- n_rel + sum(!keep) }
    }
    for (i in seq_along(wb$worksheets)) {
      ws <- wb$worksheets[[i]]
      for (f in c("drawing", "legacyDrawing", "legacyDrawingHF")) {
        v <- tryCatch(ws[[f]], error = function(e) NULL)
        if (is.character(v) && length(v)) { ws[[f]] <- character(0); n_tag <- n_tag + length(v) }
      }
    }
    if (n_rel > 0L || n_tag > 0L) {
      cat(sprintf("  已清理悬空引用：关系 %d 条、标签 %d 个\n", n_rel, n_tag))
      TRUE
    } else FALSE
  }, error = function(e) {
    cat("悬空引用清理跳过：", conditionMessage(e), "\n"); FALSE
  })
}

# 源文件是否含批注/图形部件（这类内容 openxlsx 往返会丢失，需告知用户）
source_has_rich_parts <- function(path) {
  tryCatch({
    ent <- utils::unzip(path, list = TRUE)$Name
    hits <- ent[grepl("comments|drawing|vml|media|threadedComment", ent, ignore.case = TRUE)]
    # printerSettings 是无害的打印设置，不算
    hits <- hits[!grepl("printerSettings", hits, ignore.case = TRUE)]
    length(hits) > 0L
  }, error = function(e) FALSE)
}

# 输出文件名：原名_源到目标.xlsx，并去掉非法字符
make_outname <- function(orig, src, tgt) {
  base <- sub("\\.[^.]*$", "", basename(orig))
  if (!nzchar(base)) base <- "converted"
  nm <- sprintf("%s_%s转%s.xlsx", base, src, tgt)
  nm <- gsub('[\\\\/:*?"<>|]', "_", nm)
  Encoding(nm) <- "UTF-8"
  nm
}

# ── 业务：直接输入年龄转换 ────────────────────────────────────────────────────
do_convert <- function(params) {
  src <- resolve_version(params$source)
  tgt <- resolve_version(params$target)
  if (is.null(src)) return(json_resp(list(
    error = "源版本无效", message = "source 需为版本列名，或 /versions 中给出的序号",
    hint = "GET /versions 查看全部可用版本"), 400L))
  if (is.null(tgt)) return(json_resp(list(
    error = "目标版本无效", message = "target 需为版本列名，或 /versions 中给出的序号",
    hint = "GET /versions 查看全部可用版本"), 400L))
  if (identical(src, tgt)) return(json_resp(list(
    error = "源版本与目标版本相同", message = "无需转换"), 400L))
  ages <- parse_ages(params$ages)
  if (length(ages) == 0) return(json_resp(list(
    error = "未提供有效年龄",
    hint = "ages 支持数组或逗号分隔字符串，如 251.4,250.7"), 400L))

  results <- list(); failed <- list()
  for (a in ages) {
    r <- tryCatch(convert_age(a, DF, src, tgt), error = function(e) NULL)
    if (is.null(r)) {
      failed[[length(failed) + 1]] <- list(
        input = a, reason = sprintf("无法在 %s 中定位（超出控制点范围或无相邻可用值）", src))
      next
    }
    row <- list(
      input = a, output = round(r$target, 3),
      method = if (isTRUE(r$exact)) "精确匹配" else "插值",
      stage = stage_of(a, src))
    # 插值时补充上下控制点坐标，仅供前端绘制插值数轴（不改转换算法）
    if (!isTRUE(r$exact)) {
      ex <- tryCatch({
        sx <- suppressWarnings(as.numeric(as.character(DF[[src]])))
        tx <- suppressWarnings(as.numeric(as.character(DF[[tgt]])))
        ok <- !is.na(sx) & !is.na(tx)
        lx <- sx[ok]; ly <- tx[ok]
        li <- which(lx <= a); ui <- which(lx >= a)
        if (length(li) && length(ui)) {
          li <- li[which.max(lx[li])]; ui <- ui[which.min(lx[ui])]
          list(lower = list(x = round(lx[li], 4), y = round(ly[li], 4)),
               upper = list(x = round(lx[ui], 4), y = round(ly[ui], 4)))
        } else NULL
      }, error = function(e) NULL)
      if (!is.null(ex)) { row$lower <- ex$lower; row$upper <- ex$upper }
    }
    results[[length(results) + 1]] <- row
  }
  json_resp(list(source = src, target = tgt, requested = length(ages),
                 converted = length(results), failed = length(failed),
                 results = results, errors = failed,
                 note = "只读服务：结果仅返回，不写入任何文件"))
}

# ── 业务：上传 Excel 探测结构 ─────────────────────────────────────────────────
do_inspect <- function(params) {
  b64 <- params$content %||% params$file %||% params$data
  if (is.null(b64) || !nzchar(as.character(b64))) {
    return(json_resp(list(error = "缺少文件内容", hint = "POST 字段名应为 content（base64）"), 400L))
  }
  tmp <- tryCatch(b64_to_tmp(b64), error = function(e) NULL)
  if (is.null(tmp) || !file.exists(tmp)) {
    return(json_resp(list(error = "文件解码失败", hint = "请确认上传的是 base64 编码的 .xlsx/.xls"), 400L))
  }
  on.exit(tryCatch(unlink(tmp), error = function(e) NULL), add = TRUE)

  sh <- tryCatch(readxl::excel_sheets(tmp), error = function(e) NULL)
  if (is.null(sh) || length(sh) == 0) {
    return(json_resp(list(error = "不是有效的 Excel 文件",
                          hint = "仅支持 .xlsx / .xls；若为 CSV 请先另存为 Excel"), 400L))
  }
  info <- lapply(sh, function(s) {
    d <- tryCatch(read_excel(tmp, sheet = s, n_max = 1, col_types = "text"),
                  error = function(e) NULL)
    cols <- if (is.null(d)) character(0) else names(d)
    cols <- cols[!grepl("^\\.\\.\\.$|^$", cols)]
    nrow_all <- tryCatch(nrow(read_excel(tmp, sheet = s, col_types = "text")),
                         error = function(e) NA_integer_)
    list(sheet = s, columns = as.list(cols), rows = nrow_all)
  })
  json_resp(list(ok = TRUE, sheets = info,
                 hint = "选择工作表与年龄列后调用 /convert-file"))
}

# ── 业务：上传 Excel 批量转换，返回结果文件 ───────────────────────────────────
do_convert_file <- function(params) {
  b64 <- params$content %||% params$file %||% params$data
  if (is.null(b64) || !nzchar(as.character(b64))) {
    return(json_resp(list(error = "缺少文件内容", hint = "POST 字段名应为 content（base64）"), 400L))
  }
  src <- resolve_version(params$source)
  tgt <- resolve_version(params$target)
  if (is.null(src)) return(json_resp(list(error = "源版本无效",
    hint = "source 需为版本列名或 /versions 序号"), 400L))
  if (is.null(tgt)) return(json_resp(list(error = "目标版本无效",
    hint = "target 需为版本列名或 /versions 序号"), 400L))
  if (identical(src, tgt)) return(json_resp(list(error = "源版本与目标版本相同"), 400L))

  tmp_in <- tryCatch(b64_to_tmp(b64), error = function(e) NULL)
  if (is.null(tmp_in) || !file.exists(tmp_in)) {
    return(json_resp(list(error = "文件解码失败"), 400L))
  }
  on.exit(tryCatch(unlink(tmp_in), error = function(e) NULL), add = TRUE)

  sh_all <- tryCatch(readxl::excel_sheets(tmp_in), error = function(e) NULL)
  if (is.null(sh_all) || length(sh_all) == 0) {
    return(json_resp(list(error = "不是有效的 Excel 文件"), 400L))
  }
  want_sheet <- trimws(as.character(params$sheet %||% ""))
  use_sheet <- if (nzchar(want_sheet) && want_sheet %in% sh_all) want_sheet else sh_all[1]

  df <- tryCatch(read_excel(tmp_in, sheet = use_sheet, col_types = "text"),
                 error = function(e) NULL)
  if (is.null(df) || nrow(df) == 0) {
    return(json_resp(list(error = sprintf("工作表「%s」为空或无法读取", use_sheet)), 400L))
  }

  age_col <- trimws(as.character(params$age_column %||% params$ageColumn %||% ""))
  if (!nzchar(age_col)) {
    return(json_resp(list(error = "未指定年龄列",
      hint = "先调用 /inspect-file 获取列名，再以 age_column 传入"), 400L))
  }
  if (!(age_col %in% names(df))) {
    return(json_resp(list(error = sprintf("工作表「%s」中不存在列「%s」", use_sheet, age_col),
      columns = as.list(names(df))), 400L))
  }

  # 逐行转换
  n <- nrow(df)
  raw_col <- df[[age_col]]
  out_val <- rep(NA_real_, n); out_mth <- character(n); out_stage <- character(n)
  ok <- 0L; skipped_blank <- 0L; failed <- 0L
  preview <- list()
  for (i in seq_len(n)) {
    txt <- as.character(raw_col[i])
    a <- suppressWarnings(as.numeric(txt))
    if (is.na(txt) || !nzchar(trimws(txt)) || is.na(a)) { skipped_blank <- skipped_blank + 1L; next }
    r <- tryCatch(convert_age(a, DF, src, tgt), error = function(e) NULL)
    if (is.null(r)) {
      failed <- failed + 1L
      if (length(preview) < 60) {
        preview[[length(preview) + 1]] <- list(row = i + 1L, input = a,
          output = NA, method = "无法转换", stage = "")
      }
      next
    }
    out_val[i] <- round(r$target, 3)
    out_mth[i] <- if (isTRUE(r$exact)) "精确匹配" else "插值"
    out_stage[i] <- stage_of(a, src)
    ok <- ok + 1L
    if (length(preview) < 60) {
      preview[[length(preview) + 1]] <- list(row = i + 1L, input = a,
        output = round(r$target, 3), method = out_mth[i], stage = out_stage[i])
    }
  }
  if (ok == 0L) {
    return(json_resp(list(error = "没有任何年龄转换成功",
      message = sprintf("共 %d 行，空/非数值 %d 行，超出控制点范围 %d 行", n, skipped_blank, failed),
      hint = sprintf("请确认「%s」列是数值年龄，且落在 %s 的覆盖范围内", age_col, src)), 400L))
  }

  # 写回：优先 openxlsx 保留原表样式；失败退回 writexl
  col_a <- paste0(tgt, " (来自", src, ")")
  col_b <- paste0(col_a, "_方式")
  col_c <- paste0(col_a, "_所属单元")
  tmp_out <- tempfile(fileext = ".xlsx")
  on.exit(tryCatch(unlink(tmp_out), error = function(e) NULL), add = TRUE)

  used_openxlsx <- TRUE
  sanitized <- FALSE          # 是否修正过 workbookView 空属性
  cleaned <- FALSE            # 是否清理过悬空引用
  had_rich <- source_has_rich_parts(tmp_in)   # 源文件含批注/图形，往返会丢失
  wb <- tryCatch(openxlsx::loadWorkbook(tmp_in), error = function(e) NULL)
  if (!is.null(wb)) {
    res <- tryCatch({
      # openxlsx 4.x 的正确字段是 wb$sheet_names；
      # wb$getSheetNames() 在 4.2.x 并不存在，若沿用会返回 NULL，
      # 进而把 si 兜底成 1L —— 那样无论用户选哪个表都会误写第一个表。
      sheet_names <- tryCatch(wb$sheet_names, error = function(e) NULL)
      if (is.null(sheet_names)) {
        sheet_names <- tryCatch(readxl::excel_sheets(tmp_in), error = function(e) NULL)
      }
      si <- if (is.null(sheet_names)) NA_integer_ else match(use_sheet, sheet_names)
      # 定位不到就明确失败（交给下方 writexl 兜底），绝不悄悄写到第 1 个表
      if (is.na(si)) stop(sprintf("工作表「%s」未能定位", use_sheet))
      start_col <- ncol(df) + 1L
      # 必须用 data.frame 写入：直接传字符向量会被当成「一列三行」，表头就竖排了
      new_df <- data.frame(out_val, out_mth, out_stage, stringsAsFactors = FALSE)
      names(new_df) <- c(col_a, col_b, col_c)
      openxlsx::writeData(wb, sheet = si, x = new_df,
                          startCol = start_col, startRow = 1L, colNames = TRUE)
      # 保存前修正 workbookView 空属性，保证 Excel/WPS/openpyxl 都能打开
      # 注意：tryCatch 的 {} 块在当前函数环境求值，必须用 <- 而非 <<-；
      # <<- 会去父环境找变量，找不到就在全局环境新建 —— 局部值永远不变，
      # 且多个请求会共用全局变量、互相污染。已实测复现。
      sanitized <- fix_bookviews(wb)
      # 保存前清理悬空引用，否则含批注的源文件产出后 Excel 会报「文件已损坏」
      cleaned <- clean_dangling(wb)
      openxlsx::saveWorkbook(wb, tmp_out, overwrite = TRUE)
      TRUE
    }, error = function(e) { cat("openxlsx 写回失败：", conditionMessage(e), "\n"); FALSE })
    if (!isTRUE(res)) used_openxlsx <- FALSE
  } else {
    used_openxlsx <- FALSE
  }

  if (!used_openxlsx) {
    df[[col_a]] <- out_val
    df[[col_b]] <- out_mth
    df[[col_c]] <- out_stage
    tryCatch(write_xlsx(df, tmp_out), error = function(e) {
      stop("写出结果文件失败：", conditionMessage(e))
    })
  }

  if (!file.exists(tmp_out)) return(json_resp(list(error = "结果文件生成失败"), 500L))

  orig_name <- trimws(as.character(params$filename %||% "converted.xlsx"))
  out_name <- make_outname(orig_name, src, tgt)

  # 如实告知格式保留情况：openxlsx 往返不解析批注部件，源文件若含批注则批注会丢失。
  # 填充色/列宽/合并单元格等视觉格式可保留；writexl 兜底路径则全部丢失。
  fmt_note <- if (!used_openxlsx) {
    "结果以纯数据表写出：填充色、列宽、批注等格式均未保留（数据完整）"
  } else if (isTRUE(had_rich)) {
    "已保留填充色、列宽等格式；但源文件中的批注/图形在转换中无法保留（数据完整）"
  } else {
    "已保留源文件的填充色、列宽等格式"
  }

  json_resp(list(
    ok = TRUE,
    sheet = use_sheet,
    age_column = age_col,
    source = src, target = tgt,
    rows = n, converted = ok, failed = failed, skipped = skipped_blank,
    format_preserved = used_openxlsx,
    comments_lost = isTRUE(had_rich) && used_openxlsx,
    dangling_cleaned = isTRUE(cleaned),
    ooxml_sanitized = isTRUE(sanitized),
    format_note = fmt_note,
    new_columns = list(col_a, col_b, col_c),
    preview = preview,
    filename = out_name,
    content = tmp_to_b64(tmp_out),
    note = "结果以下载形式返回，服务未修改你上传的原文件"
  ))
}

# ── 路由 ──────────────────────────────────────────────────────────────────────
route <- function(req) {
  method <- toupper(req$REQUEST_METHOD %||% "GET")
  path <- req$PATH_INFO
  if (is.null(path) || !nzchar(path)) path <- "/"

  if (method == "OPTIONS") {
    return(list(status = 204L, headers = list(
      "Access-Control-Allow-Origin" = "*",
      "Access-Control-Allow-Methods" = "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers" = "Content-Type",
      "Access-Control-Allow-Max-Age" = "86400"
    ), body = ""))
  }

  if (path == "/shutdown") {
    resp <- json_resp(list(ok = TRUE, message = "服务已关闭，正在退出"))
    tryCatch(httpuv::stopServer(), error = function(e) NULL)
    return(resp)
  }

  # 静态托管
  safe <- gsub("^/+", "", path)
  safe <- gsub("\\.\\.", "", safe)               # 防目录穿越
  static_file <- if (safe == "" || safe == "index.html") {
    file.path(WEB_DIR, "index.html")
  } else {
    file.path(WEB_DIR, safe)
  }
  if (file.exists(static_file) && !dir.exists(static_file)) {
    ext <- tolower(sub(".*\\.", "", basename(static_file)))
    ctype <- switch(ext,
      html = "text/html; charset=utf-8",
      css = "text/css; charset=utf-8",
      js = "application/javascript; charset=utf-8",
      json = "application/json; charset=utf-8",
      svg = "image/svg+xml",
      "text/plain; charset=utf-8")
    bytes <- readBin(static_file, what = "raw", n = file.info(static_file)$size)
    return(list(status = 200L,
                headers = list("Content-Type" = ctype,
                               "Access-Control-Allow-Origin" = "*",
                               "Cache-Control" = "no-store"),
                body = bytes))
  }

  params <- parse_qs(req$QUERY_STRING)
  if (method == "POST") {
    raw <- tryCatch(req$rook.input$read(), error = function(e) NULL)
    if (is.raw(raw)) {
      if (length(raw) > MAX_BYTES) {
        return(json_resp(list(error = "上传内容过大",
          hint = sprintf("上限约 %d MB", as.integer(MAX_BYTES / 1024 / 1024))), 413L))
      }
      txt <- rawToChar(raw)
      Encoding(txt) <- "UTF-8"
      raw <- txt
    }
    if (!is.null(raw) && nzchar(raw)) {
      body <- tryCatch(fromJSON(raw, simplifyVector = TRUE), error = function(e) NULL)
      if (is.null(body)) return(json_resp(list(error = "请求体不是合法 JSON"), 400L))
      if (is.list(body)) for (k in names(body)) params[[k]] <- body[[k]]
    }
  }

  if (path %in% c("/api", "/help")) {
    return(json_resp(list(
      service = "年代表年龄转换 API（文件上传版）",
      说明 = "把各版本地质年代表中对应地层界线作为控制点，将历史年龄由源版本转换到目标版本",
      只读 = TRUE,
      endpoints = list(
        "GET  /health" = "健康检查",
        "GET  /versions" = "可用版本列表（序号可直接作为 source/target）",
        "GET  /convert" = "?source=ICS 200404&target=ICS 202606&ages=251.4,250.7",
        "POST /convert" = 'body: {"source":..,"target":..,"ages":[251.4,250.7]}',
        "POST /inspect-file" = 'body: {"filename":"a.xlsx","content":"<base64>"} -> 工作表与列名',
        "POST /convert-file" = 'body: {"filename","content","sheet","age_column","source","target"} -> 结果文件 base64'
      ),
      参数 = list(
        source = "源版本：列名或 /versions 中的序号",
        target = "目标版本：列名或序号",
        ages = "待转换年龄：数组，或逗号/空格分隔字符串",
        content = "上传文件的 base64（可带 data:...;base64, 前缀）",
        age_column = "Excel 中年龄所在列名（由 /inspect-file 获得）"
      ),
      安全 = "上传文件仅写入系统临时目录并在响应后删除；绝不修改或覆盖用户原文件"
    )))
  }

  if (path == "/health") {
    return(json_resp(list(status = "ok", rows = nrow(DF), cols = ncol(DF),
      versions = length(CANDIDATES), readonly = TRUE, upload = TRUE)))
  }
  # 页面心跳：网页每 5 秒上报一次，用于判断标签页是否已关闭
  if (path == "/ping") {
    LAST_PING <<- Sys.time()
    GOT_PING <<- TRUE
    return(json_resp(list(pong = TRUE)))
  }
  if (path == "/versions") {
    vers <- lapply(seq_along(CANDIDATES), function(i) {
      col <- suppressWarnings(as.numeric(as.character(DF[[CANDIDATES[i]]])))
      list(index = i, name = CANDIDATES[i], cover = sum(!is.na(col)))
    })
    return(json_resp(list(count = length(CANDIDATES), versions = vers)))
  }
  if (path == "/convert") {
    if (!(method %in% c("GET", "POST"))) {
      return(json_resp(list(error = "该接口仅支持 GET / POST"), 405L))
    }
    return(do_convert(params))
  }
  if (path == "/inspect-file") {
    if (method != "POST") return(json_resp(list(error = "该接口仅支持 POST"), 405L))
    return(do_inspect(params))
  }
  if (path == "/convert-file") {
    if (method != "POST") return(json_resp(list(error = "该接口仅支持 POST"), 405L))
    return(do_convert_file(params))
  }

  json_resp(list(error = "未知接口", path = path, hint = "GET /api 查看接口说明"), 404L)
}

app <- list(call = function(req) {
  tryCatch(route(req), error = function(e) {
    json_resp(list(error = "服务器内部错误", detail = conditionMessage(e)), 500L)
  })
})

# ── 启动 ──────────────────────────────────────────────────────────────────────
# 出错时不要立刻退出：双击启动的窗口一闪而过就什么也看不到了。
# 打印原因并等待按键，让用户能看到到底哪一步失败。
pause_and_exit <- function(status = 1L) {
  cat("\n")
  cat(strrep("-", 66), "\n")
  cat("启动未成功。请截图上面的错误信息。\n")
  cat("按回车键关闭本窗口…")
  invisible(tryCatch(readLines(stdin(), n = 1), error = function(e) NULL))
  quit(status = status)
}

# 端口被占用时自动往后找（双击启动最常见的失败原因就是端口已被占用）
find_free_port <- function(host, first_port) {
  tried <- integer(0)
  for (p in seq.int(first_port, first_port + 19L)) {
    tried <- c(tried, p)
    s <- tryCatch(httpuv::startServer(host, p, app), error = function(e) NULL)
    if (!is.null(s)) return(list(port = p, srv = s, tried = tried))
  }
  list(port = NA_integer_, srv = NULL, tried = tried)
}

cat("\n")
cat(strrep("=", 66), "\n")
cat("  年代表年龄转换（文件上传版）\n")
cat("  接口： /  /health  /versions  /convert  /inspect-file  /convert-file\n")
cat("  模式： 只读（结果以下载返回，不修改你的原文件）\n")
cat(sprintf("  数据：%s\n", basename(XLSX)))
cat(strrep("=", 66), "\n")

found <- find_free_port(HOST, PORT)
if (is.null(found$srv)) {
  cat(sprintf("\n端口 %d~%d 全部被占用，无法启动。\n", min(found$tried), max(found$tried)))
  cat("可能已有一个同样的服务在运行：先关掉之前弹出的窗口，再重新双击。\n")
  cat("或者在命令后加 --port=8800 指定别的端口。\n")
  pause_and_exit(1L)
}
srv <- found$srv
PORT <- found$port
if (PORT != found$tried[1]) {
  cat(sprintf("\n注意：原定端口 %d 已被占用，本次改用端口 %d\n", found$tried[1], PORT))
}
url <- sprintf("http://%s:%d/", ifelse(HOST %in% c("0.0.0.0", "*"), "127.0.0.1", HOST), PORT)
cat(sprintf("\n  监听： %s\n", url))

# 自动打开浏览器：依次尝试 shell(start)、browseURL、直接调 Edge/Chrome。
# 任一成功即停止；全部失败则提示手动打开（服务仍在运行，不影响使用）。
opened <- FALSE
if (.Platform$OS.type == "windows") {
  opened <- isTRUE(tryCatch({ shell(paste0('start "" "', url, '"')); TRUE },
                            error = function(e) FALSE, warning = function(w) FALSE))
} else {
  opened <- isTRUE(tryCatch({ system2("xdg-open", url, wait = FALSE); TRUE },
                            error = function(e) FALSE))
}
if (!opened) {
  opened <- isTRUE(tryCatch({ utils::browseURL(url); TRUE }, error = function(e) FALSE))
}
if (!opened) {
  browsers <- c(
    file.path(Sys.getenv("PROGRAMFILES"), "Microsoft\\Edge\\Application\\msedge.exe"),
    file.path(Sys.getenv("PROGRAMFILES(X86)"), "Microsoft\\Edge\\Application\\msedge.exe"),
    file.path(Sys.getenv("LOCALAPPDATA"), "Google\\Chrome\\Application\\chrome.exe"),
    file.path(Sys.getenv("PROGRAMFILES"), "Google\\Chrome\\Application\\chrome.exe"))
  for (b in browsers[nzchar(browsers)]) {
    if (file.exists(b)) {
      opened <- isTRUE(tryCatch({ system2(b, shQuote(url), wait = FALSE); TRUE },
                                error = function(e) FALSE))
      if (opened) break
    }
  }
}

cat("\n")
cat(strrep("=", 66), "\n")
if (opened) {
  cat("  已尝试在默认浏览器中打开。\n")
} else {
  cat("  ★ 浏览器未能自动打开，请手动复制下面这行地址到浏览器：\n")
}
cat(sprintf("      %s\n", url))
cat(strrep("=", 66), "\n")
cat("  本窗口不要关闭（关了服务就停了）；停止服务按 Ctrl+C\n\n")

# 浏览器标签页关闭时自动退出：
# 网页每 5 秒发一次 /ping，收到首个心跳后，超过 90 秒没再收到即判定页面已关闭。
# （不能靠服务自己探测 /health——自己请求自己永远成功，检测不到页面关闭）
# 若浏览器始终没打开过（无心跳），则不自杀，避免刚启动就退出。
LAST_PING <- Sys.time()
GOT_PING <- FALSE
IDLE_LIMIT <- 90
on.exit(tryCatch({ httpuv::stopServer(srv); cat("服务已停止\n") }, error = function(e) NULL), add = TRUE)
repeat {
  httpuv::service(0.2)
  Sys.sleep(0.2)
  if (GOT_PING && as.numeric(difftime(Sys.time(), LAST_PING, units = "secs")) > IDLE_LIMIT) {
    cat(sprintf("\n已 %d 秒未收到网页心跳，判定页面已关闭，服务自动退出\n", IDLE_LIMIT))
    break
  }
}
