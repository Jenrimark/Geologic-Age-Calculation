#!/usr/bin/env Rscript
# ═══════════════════════════════════════════════════════════════════════════
#  年代表年龄转换 · API 服务
#  把「各版本直接转换_无中介.R」的转换能力封装为 HTTP / JSON 接口
#
#  与命令行版的三点关键差异（前两点是性能与形态，第三点是安全）：
#    1) 输入改为 HTTP 参数，不再 readline 交互
#    2) 年代表启动时读取一次、常驻内存，不再每次请求重读磁盘
#    3) ★只读服务：绝不写回任何文件★
#       命令行版结尾的 write_xlsx(src_df, file_path) 会覆盖用户原表；
#       挂到网络后任何人都能借一个请求覆盖你的数据，并发写还会损坏文件，
#       因此 API 模式彻底移除写回，只返回 JSON 结果。
#
#  转换逻辑不重复实现：启动时从原脚本载入 find_age_in_column / convert_age，
#  原脚本改了算法，本服务重启即同步生效。
#
#  启动： Rscript 年代表转换_API服务.R
#  参数： --port=8787   --host=127.0.0.1
#        --xlsx=年代表路径（缺省自动在脚本同目录 / 环境变量 GEO_XLSX / 桌面 WJY2 查找）
#        --core=核心脚本路径（缺省取同目录「各版本直接转换_无中介.R」）
#  停止： Ctrl + C
#
#  接口：
#    GET  /                     接口说明
#    GET  /health               健康检查
#    GET  /versions             可用版本列表（含序号，序号可直接当 source/target）
#    GET  /convert?source=&target=&ages=251.4,250.7
#    POST /convert              body: {"source":"GTS2004","target":"GTS2020","ages":[251.4,250.7]}
# ═══════════════════════════════════════════════════════════════════════════
suppressPackageStartupMessages({
  library(httpuv)
  library(jsonlite)
  library(readxl)
})
# ── 配置（全部支持命令行参数覆盖，跨平台自适应，兼容 Windows/macOS/不同用户） ──
MAX_AGES <- 10000          # 单次请求最多接受的年龄个数，防止超量请求拖死服务

args <- commandArgs(trailingOnly = TRUE)
get_arg <- function(key, default) {
  hit <- grep(paste0("^--", key, "="), args, value = TRUE)
  if (length(hit)) sub(paste0("^--", key, "="), "", hit[1]) else default
}

# 脚本自身目录：Rscript 经 --file= 传入；交互式回退工作目录
SCRIPT_DIR <- (function() {
  fa <- commandArgs(trailingOnly = FALSE)
  hit <- fa[grepl("^--file=", fa)]
  if (length(hit)) dirname(normalizePath(sub("^--file=", "", hit[1]))) else getwd()
})()

# 核心转换脚本：--core= > 与本服务同目录的「各版本直接转换_无中介.R」
USER_R <- get_arg("core", file.path(SCRIPT_DIR, "各版本直接转换_无中介.R"))

# 年代表：--xlsx= > 环境变量 GEO_XLSX > 同目录候选 > 用户桌面 WJY2 回退
resolve_xlsx <- function() {
  candidates <- c(
    get_arg("xlsx"),
    Sys.getenv("GEO_XLSX", unset = ""),
    file.path(SCRIPT_DIR, "地质年代表（表头）.xlsx"),
    file.path(SCRIPT_DIR, "wjy.xlsx"),
    file.path(path.expand("~"), "Desktop", "WJY2", "地质年代表（表头）.xlsx"),
    file.path(path.expand("~"), "Desktop", "WJY2", "wjy.xlsx")
  )
  candidates <- candidates[nzchar(candidates)]
  hit <- candidates[file.exists(candidates)]
  if (length(hit)) return(normalizePath(hit[1]))
  cat("未找到年代表，已按以下顺序查找：\n")
  for (p in candidates) cat("  -", p, "\n")
  stop("请用 --xlsx=路径 指定年代表，或设置环境变量 GEO_XLSX")
}
XLSX  <- resolve_xlsx()
SHEET <- "Sheet1"

PORT <- suppressWarnings(as.integer(get_arg("port", "8787")))
HOST <- get_arg("host", "127.0.0.1")
if (is.na(PORT)) PORT <- 8787L
# ── 载入原脚本的转换逻辑（剥离 library() 与末尾 main() 调用，避免触发交互） ──
load_core <- function(path) {
  if (!file.exists(path)) stop("找不到核心脚本：", path)
  lines <- readLines(path, encoding = "UTF-8", warn = FALSE)
  keep  <- !grepl("^\\s*library\\(", lines) & !grepl("^\\s*main\\(\\)\\s*$", lines)
  env   <- new.env(parent = globalenv())
  eval(parse(text = paste(lines[keep], collapse = "\n")), envir = env)
  for (fn in c("find_age_in_column", "convert_age")) {
    if (!is.function(get0(fn, envir = env))) {
      stop("核心脚本中未找到函数 ", fn, "，无法提供转换服务")
    }
  }
  env
}
cat("载入转换逻辑：", USER_R, "\n")
CORE        <- load_core(USER_R)
convert_age <- CORE$convert_age
# ── 读取年代表（启动时一次，常驻内存） ──────────────────────────────────────
cat("读取年代表：", XLSX, "\n", sep = "")
# sheet 自适应：优先 Sheet1，不存在则取第一个工作表
sheets <- readxl::excel_sheets(XLSX)
use_sheet <- if (SHEET %in% sheets) SHEET else sheets[1]
DF <- read_excel(XLSX, sheet = use_sheet, col_types = "text")
cat(sprintf("  %d 行 x %d 列\n", nrow(DF), ncol(DF)))
# 可用版本列：排除误差列、空列名，以及宇/界/系/统/阶五列地层层级列
# （原命令行脚本读的是另一份表，其列布局不含这五列；表头版前 5 列为层级标识，须排除）
HIER_COLS <- c("Eonothem", "Erathem", "System", "Series", "Stage",
               "宇", "界", "系", "统", "阶")
ALL_COLS   <- names(DF)
CANDIDATES <- ALL_COLS[!grepl("误差|error|Unnamed", ALL_COLS, ignore.case = TRUE)]
CANDIDATES <- CANDIDATES[CANDIDATES != ""]
CANDIDATES <- CANDIDATES[!(CANDIDATES %in% HIER_COLS)]
cat(sprintf("  可用版本 %d 个\n", length(CANDIDATES)))
# ── 工具函数 ────────────────────────────────────────────────────────────────
# URL 百分号解码（httpuv 不解码 QUERY_STRING；版本名含空格与中文，必须自行解码）
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
# 解析查询串为命名 list
# 注意：httpuv 的 QUERY_STRING 含前导「?」（如 "?source=GTS2004&target=GTS2020"），
# 必须先剥掉，否则首个参数名会变成「?source」而取不到值。
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
# R 4.x 自带 %||%，旧版没有，这里兜底定义（route() 中使用）
`%||%` <- function(a, b) if (is.null(a)) b else a
json_resp <- function(obj, status = 200L) {
  list(
    status  = status,
    headers = list(
      "Content-Type"              = "application/json; charset=utf-8",
      "Access-Control-Allow-Origin" = "*",      # 允许网页跨域调用
      "Cache-Control"             = "no-store"
    ),
    body = toJSON(obj, auto_unbox = TRUE, null = "null", na = "null")
  )
}
# 版本名解析：接受「序号」或「列名」两种写法
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
# 就近匹配源列，定位该年龄所属地层单元（用于结果核查）
stage_of <- function(a, src_col) {
  vals <- suppressWarnings(as.numeric(as.character(DF[[src_col]])))
  if (all(is.na(vals))) return(NA_character_)
  idx <- which.min(abs(vals - a))
  for (cn in c("阶", "Stage")) {
    if (cn %in% names(DF)) {
      v <- DF[[cn]][idx]
      if (!is.na(v) && nzchar(trimws(as.character(v)))) return(as.character(v))
    }
  }
  NA_character_
}
# 解析年龄列表：支持数组或「逗号/空格/分号/顿号」分隔字符串
parse_ages <- function(x) {
  if (is.null(x)) return(numeric(0))
  if (is.list(x)) x <- unlist(x, use.names = FALSE)
  if (is.numeric(x)) v <- x else {
    tok <- unlist(strsplit(paste(as.character(x), collapse = " "),
                           "[,，;；、[:space:]]+"))
    tok <- tok[tok != ""]
    v <- suppressWarnings(as.numeric(tok))
  }
  v <- v[!is.na(v)]
  v[is.finite(v)]
}
# ── 业务：执行转换 ──────────────────────────────────────────────────────────
do_convert <- function(params) {
  src <- resolve_version(params$source)
  dst <- resolve_version(params$target)
  if (is.null(src) || is.null(dst)) {
    return(json_resp(list(
      error   = "版本参数无效",
      message = "source / target 需为版本列名，或 /versions 中给出的序号",
      hint    = "GET /versions 查看全部可用版本"
    ), 400L))
  }
  if (identical(src, dst)) {
    return(json_resp(list(error = "源版本与目标版本相同，无需转换",
                          source = src, target = dst), 400L))
  }
  ages <- parse_ages(params$ages)
  if (length(ages) == 0) {
    return(json_resp(list(error = "未提供有效年龄",
                          hint  = "ages 支持数组或逗号分隔字符串，如 251.4,250.7"), 400L))
  }
  if (length(ages) > MAX_AGES) {
    return(json_resp(list(error = sprintf("单次最多 %d 个年龄，当前 %d 个",
                                          MAX_AGES, length(ages))), 400L))
  }
  results <- list(); failed <- list()
  for (a in ages) {
    r <- convert_age(a, DF, src, dst)
    if (is.null(r)) {
      failed <- c(failed, list(list(input = a,
        reason = sprintf("无法在 %s 中定位（超出控制点范围或无相邻可用值）", src))))
    } else {
      results <- c(results, list(list(
        input  = r$source,
        output = round(r$target, 4),
        method = if (isTRUE(r$exact)) "精确匹配" else "插值",
        stage  = stage_of(a, src)
      )))
    }
  }
  json_resp(list(
    source    = src,
    target    = dst,
    requested = length(ages),
    converted = length(results),
    failed    = length(failed),
    results   = results,
    errors    = failed,
    note      = "只读服务：结果仅返回，不写入任何文件"
  ))
}
# ── 路由 ────────────────────────────────────────────────────────────────────
route <- function(req) {
  method <- toupper(req$REQUEST_METHOD %||% "GET")
  path   <- req$PATH_INFO
  if (is.null(path) || !nzchar(path)) path <- "/"
  # CORS 预检
  if (method == "OPTIONS") {
    return(list(status = 204L, headers = list(
      "Access-Control-Allow-Origin"  = "*",
      "Access-Control-Allow-Methods" = "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers" = "Content-Type",
      "Access-Control-Max-Age"       = "86400"
    ), body = ""))
  }
  # 取参数：GET 用查询串，POST 用 JSON body（两者都支持，POST 优先）
  params <- parse_qs(req$QUERY_STRING)
  if (method == "POST") {
    raw <- tryCatch(req$rook.input$read(), error = function(e) NULL)
    if (is.raw(raw)) {                 # 二进制体 -> 字符并标记 UTF-8（中文列名/年龄需要）
      txt <- rawToChar(raw)
      Encoding(txt) <- "UTF-8"
      raw <- txt
    }
    if (!is.null(raw) && nzchar(raw)) {
      body <- tryCatch(fromJSON(raw, simplifyVector = TRUE),
                       error = function(e) NULL)
      if (is.null(body)) {
        return(json_resp(list(error = "请求体不是合法 JSON"), 400L))
      }
      if (is.list(body)) for (k in names(body)) params[[k]] <- body[[k]]
    }
  }
  if (path %in% c("/", "/index", "/api")) {
    return(json_resp(list(
      service = "年代表年龄转换 API",
      说明     = "把各版本地质年代表中对应地层界线作为控制点，将历史年龄由源版本转换到目标版本",
      只读     = TRUE,
      endpoints = list(
        "GET /health"    = "健康检查",
        "GET /versions"  = "可用版本列表（序号可直接作为 source/target）",
        "GET /convert"   = "?source=GTS2004&target=GTS2020&ages=251.4,250.7",
        "POST /convert"  = 'body: {"source":"GTS2004","target":"GTS2020","ages":[251.4,250.7]}'
      ),
      参数 = list(
        source = "源版本：列名或 /versions 中的序号",
        target = "目标版本：列名或序号",
        ages   = "待转换年龄：数组，或逗号/空格分隔字符串"
      ),
      返回字段 = list(
        output = "转换后年龄(Ma)",
        method = "精确匹配 或 插值",
        stage  = "就近匹配到的所属地层单元，便于核查"
      )
    )))
  }
  if (path == "/health") {
    return(json_resp(list(
      status = "ok", rows = nrow(DF), cols = ncol(DF),
      versions = length(CANDIDATES), readonly = TRUE
    )))
  }
  if (path == "/versions") {
    return(json_resp(list(
      count = length(CANDIDATES),
      versions = lapply(seq_along(CANDIDATES), function(i)
        list(index = i, name = CANDIDATES[i]))
    )))
  }
  if (path == "/convert") {
    if (!(method %in% c("GET", "POST"))) {
      return(json_resp(list(error = "该接口仅支持 GET / POST"), 405L))
    }
    return(do_convert(params))
  }
  json_resp(list(error = "未知接口", path = path, hint = "GET / 查看接口说明"), 404L)
}
app <- list(
  call = function(req) {
    tryCatch(route(req), error = function(e) {
      json_resp(list(error = "服务器内部错误", detail = conditionMessage(e)), 500L)
    })
  }
)
# ── 启动 ────────────────────────────────────────────────────────────────────
cat("\n", strrep("=", 66), "\n", sep = "")
cat(sprintf("  年代表年龄转换 API 已启动\n"))
cat(sprintf("  监听： http://%s:%d\n", HOST, PORT))
cat(sprintf("  接口： /  /health  /versions  /convert\n"))
cat(sprintf("  模式： 只读（不写回任何文件）\n"))
if (HOST == "0.0.0.0") {
  cat("  ⚠ 已绑定全部网卡：同一网络内的其他设备均可访问本服务。\n")
  cat("    本服务无身份验证，请勿在不可信网络中长期开启。\n")
}
cat(strrep("=", 66), "\n")
cat("  按 Ctrl+C 停止\n\n")
httpuv::runServer(HOST, PORT, app)
