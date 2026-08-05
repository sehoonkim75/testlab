# ==============================================================================
# CTT 문항분석기 — Shiny GUI (R 기반) v9
# ------------------------------------------------------------------------------
# v9에서 바뀐 점:
#   1) PDF 다운로드 기능 제거 — Chrome/Edge 헤드리스 변환이 환경에 따라 불안정해
#      아예 뺐습니다. 대신 "HTML 보고서 다운로드"만 제공합니다. 받은 HTML을
#      브라우저에서 열어 인쇄(Ctrl+P → PDF로 저장)하면 동일한 결과를 얻을 수
#      있습니다 — 어떤 환경에서도 항상 되는 방법이라 이쪽이 더 안정적입니다.
#   2) 영역내 문항-총점상관 추가 — 영역(과목) 구성이 이질적인 시험에서는 전체
#      총점보다 "그 영역 안에서의" 총점과의 상관이 더 적절한 변별도 지표일 수
#      있어, 문항분석표에 전체상관과 영역상관을 나란히 표시합니다.
#   3) 선지별 반응분포 표에도 영역 열 추가 — 화면 표시, 엑셀, HTML 보고서 모두
#      반영했습니다.
#
# ------------------------------------------------------------------------------
# 실행 방법
# ------------------------------------------------------------------------------
#   1) 최초 1회만 패키지 설치:
#        install.packages(c("shiny","DT","readxl","writexl","CTT",
#                            "showtext","sysfonts","base64enc"))
#   2) RStudio에서 이 파일(app.R)을 열고 오른쪽 위 "Run App" 버튼 클릭
#      (또는 R 콘솔에서: shiny::runApp("app.R"))
#   3) 동료와 함께 쓰려면 https://www.shinyapps.io 배포 또는 기관 Shiny Server 이용
#
# ------------------------------------------------------------------------------
# 정답표 파일 작성법 (영역 구분 포함)
# ------------------------------------------------------------------------------
#   가로형: 1행 문항명 / 2행 정답 / 3행 영역(선택)
#   세로형: 문항, 정답, 영역(선택) 3개 열
# ==============================================================================

library(shiny)
library(DT)
library(readxl)
library(writexl)
library(CTT)
library(sysfonts)
library(showtext)
library(base64enc)

# ------------------------------------------------------------------------------
# 판정기준표 (참고용 — 문항분석표에는 자동 적용하지 않음)
# ------------------------------------------------------------------------------
CRITERIA_TABLE <- data.frame(
  `문항-총점상관(rit)` = c(".40 이상", ".30 ~ .39", ".20 ~ .29", ".20 미만"),
  해석 = c("매우 좋음", "좋음", "검토 필요", "제거 검토"),
  설명 = c(
    "상위권과 하위권 응시자를 뚜렷하게 구분하는 문항",
    "무난하게 상/하위 응시자를 구분하는 문항",
    "구분력이 약한 편이라 수정 검토가 필요한 문항",
    "구분력이 거의 없거나 음(-)의 상관으로, 제외를 검토할 문항"
  ),
  stringsAsFactors = FALSE, check.names = FALSE
)

# ------------------------------------------------------------------------------
# 한글 폰트 등록 (그래프 PNG용)
# ------------------------------------------------------------------------------
register_korean_font <- function(){
  candidates <- c(
    "C:/Windows/Fonts/malgun.ttf",
    "C:/Windows/Fonts/malgunbd.ttf",
    "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
    "/Library/Fonts/AppleGothic.ttf",
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"
  )
  found <- candidates[file.exists(candidates)]
  if(length(found) > 0){
    sysfonts::font_add(family = "malgun", regular = found[1])
    showtext::showtext_auto()
    showtext::showtext_opts(dpi = 130)
    return("malgun")
  }
  NA_character_
}
KOR_FONT <- register_korean_font()

# ------------------------------------------------------------------------------
# 데이터 읽기 / 정답표 파싱
# ------------------------------------------------------------------------------
read_table_any <- function(path, ext, sheet = NULL){
  if(ext %in% c("xlsx", "xls")){
    as.data.frame(read_excel(path, sheet = sheet, col_names = TRUE))
  } else {
    read.csv(path, header = TRUE, stringsAsFactors = FALSE, check.names = FALSE)
  }
}

parse_key_and_domain <- function(key_df, item_names){
  header <- colnames(key_df)
  domain_vec <- setNames(rep(NA_character_, length(item_names)), item_names)

  if(any(header %in% item_names)){
    key_vec <- as.character(unlist(key_df[1, item_names]))
    names(key_vec) <- item_names
    if(nrow(key_df) >= 2){
      dom_row <- as.character(unlist(key_df[2, item_names]))
      domain_vec[item_names] <- dom_row
    }
  } else {
    key_vec <- setNames(as.character(key_df[[2]]), as.character(key_df[[1]]))
    key_vec <- key_vec[item_names]
    dom_col <- which(grepl("영역|domain", header, ignore.case = TRUE))
    if(length(dom_col) == 0 && ncol(key_df) >= 3) dom_col <- 3
    if(length(dom_col) >= 1){
      dom_map <- setNames(as.character(key_df[[dom_col[1]]]), as.character(key_df[[1]]))
      domain_vec[item_names] <- dom_map[item_names]
    }
  }

  missing <- item_names[is.na(key_vec) | key_vec == ""]
  if(length(missing) > 0) stop("정답표에 다음 문항의 정답이 없습니다: ", paste(missing, collapse = ", "))

  domain_vec[is.na(domain_vec) | domain_vec == ""] <- "미지정"
  list(key = key_vec, domain = domain_vec)
}

pivot_distractor_wide <- function(dlong, item_names, choice_values, domain_vec){
  rows <- lapply(item_names, function(nm){
    sub <- dlong[dlong$문항 == nm, ]
    correct_row <- sub[!is.na(sub$correct) & trimws(as.character(sub$correct)) == "*", ]
    correct_opt <- if(nrow(correct_row) > 0) correct_row$선지[1] else NA_character_

    pct_list <- lapply(choice_values, function(opt){
      r <- sub[sub$선지 == as.character(opt), ]
      if(nrow(r) == 0) return(NA_real_)
      round(r$rspP[1] * 100)
    })
    names(pct_list) <- paste0("선지", choice_values, "(%)")

    flagged_opts <- sub$선지[!is.na(sub$역작동의심) & sub$역작동의심 == "Y"]
    flagged_str <- if(length(flagged_opts) > 0) paste(flagged_opts, collapse = ",") else "-"

    row <- c(list(문항 = nm, 영역 = unname(domain_vec[[nm]])), list(정답 = correct_opt), pct_list, list(역작동의심 = flagged_str))
    as.data.frame(row, stringsAsFactors = FALSE, check.names = FALSE)
  })
  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

# ------------------------------------------------------------------------------
# 채점 및 CTT 분석
# ------------------------------------------------------------------------------
run_ctt <- function(raw_df, key_df, has_id, n_groups){
  if(has_id){
    item_names <- colnames(raw_df)[-1]
    items <- raw_df[, -1, drop = FALSE]
  } else {
    item_names <- colnames(raw_df)
    items <- raw_df
  }
  if(length(item_names) < 2) stop("문항이 2개 미만입니다. 최소 3문항을 권장합니다.")

  parsed <- parse_key_and_domain(key_df, item_names)
  key_vec <- parsed$key
  domain_vec <- parsed$domain

  items <- as.data.frame(lapply(items, as.character), stringsAsFactors = FALSE)
  colnames(items) <- item_names

  scored <- score(items, key_vec, output.scored = TRUE)
  ia <- itemAnalysis(scored$scored, itemReport = TRUE)

  # 영역별로 "그 영역 안에서의" 문항-총점상관을 별도 계산 (CTT::itemAnalysis를 영역 부분집합에 재적용)
  # 영역 구성이 이질적일수록, 전체총점보다 이 값이 더 적절한 변별도 지표일 수 있습니다.
  domains <- unique(domain_vec[item_names])
  item_domain_corr <- setNames(rep(NA_real_, length(item_names)), item_names)
  domain_rows <- list()
  for(dm in domains){
    idx <- which(domain_vec[item_names] == dm)
    sub_items <- scored$scored[, idx, drop = FALSE]
    n_it <- length(idx)
    dom_alpha <- NA_real_
    if(n_it >= 2){
      sub_ia <- tryCatch(itemAnalysis(sub_items, itemReport = TRUE), error = function(e) NULL)
      if(!is.null(sub_ia)){
        dom_alpha <- sub_ia$alpha
        item_domain_corr[item_names[idx]] <- round(sub_ia$itemReport$pBis, 2)
      }
    }
    dom_total <- if(n_it >= 1) rowSums(sub_items) else rep(0, nrow(sub_items))
    domain_rows[[dm]] <- data.frame(
      영역 = dm, 문항수 = n_it,
      `평균정답률` = round(mean(ia$itemReport$itemMean[idx]), 2),
      `영역신뢰도(alpha)` = round(dom_alpha, 2),
      `평균영역상관` = round(mean(item_domain_corr[item_names[idx]], na.rm = TRUE), 2),
      영역총점평균 = round(mean(dom_total), 2),
      영역총점SD = round(sd(dom_total), 2),
      stringsAsFactors = FALSE, check.names = FALSE
    )
  }
  domain_summary <- do.call(rbind, domain_rows)
  domain_summary <- domain_summary[order(-domain_summary$문항수), ]

  item_stats <- data.frame(
    문항 = item_names,
    영역 = domain_vec[item_names],
    응답수 = nrow(scored$scored),
    평균 = round(ia$itemReport$itemMean, 2),
    `전체상관` = round(ia$itemReport$pBis, 2),
    `영역상관` = unname(item_domain_corr[item_names]),
    삭제시알파 = round(ia$itemReport$alphaIfDeleted, 2),
    stringsAsFactors = FALSE, check.names = FALSE
  )

  total_scores <- scored$score
  summary_df <- data.frame(
    항목 = c("응시자 수", "문항 수", "Cronbach's alpha", "총점 평균", "총점 표준편차", "총점 범위"),
    값 = c(nrow(scored$scored), length(item_names), round(ia$alpha, 2),
           round(mean(total_scores), 2), round(sd(total_scores), 2),
           paste(min(total_scores), "-", max(total_scores))),
    stringsAsFactors = FALSE
  )

  distractor_wide <- NULL
  choice_values <- sort(unique(suppressWarnings(as.numeric(unlist(items)))))
  choice_values <- choice_values[!is.na(choice_values)]

  da <- tryCatch(distractorAnalysis(items, key_vec, nGroups = n_groups), error = function(e) NULL)
  if(!is.null(da) && length(choice_values) > 0){
    dlong <- do.call(rbind, lapply(seq_along(da), function(i){
      d <- as.data.frame(da[[i]])
      d$문항 <- item_names[i]
      d$선지 <- rownames(d)
      rownames(d) <- NULL
      d
    }))
    is_correct_row <- !is.na(dlong$correct) & trimws(as.character(dlong$correct)) == "*"
    dlong$역작동의심 <- ifelse(!is_correct_row & dlong$upper > dlong$lower, "Y", "")

    distractor_wide <- pivot_distractor_wide(dlong, item_names, choice_values, domain_vec)
  }

  list(item_stats = item_stats, summary_df = summary_df, distractor_df = distractor_wide,
       choice_values = choice_values, domain_summary = domain_summary,
       total_scores = total_scores, ia = ia, item_names = item_names)
}

# 화면(Shiny)·리포트 공용 그래프 함수
make_scatter <- function(res, main_title = "문항 난이도-변별도 산점도"){
  if(!is.na(KOR_FONT)) par(family = KOR_FONT)
  disc <- res$ia$itemReport$pBis
  p    <- res$ia$itemReport$itemMean
  col  <- ifelse(disc < 0.2, "#A8402C", ifelse(disc < 0.3, "#B9862F", "#2F4A6B"))
  plot(p, disc, pch = 19, col = col, cex = 1.8, xlim = c(0, 1), ylim = c(-0.2, 1),
       xlab = "난이도 (p)", ylab = "문항-총점 상관", main = main_title)
  text(p, disc, labels = res$item_names, pos = 3, cex = 0.8)
  abline(h = 0.2, lty = 2, col = "grey70")
}
make_hist <- function(res, main_title = "총점 분포"){
  if(!is.na(KOR_FONT)) par(family = KOR_FONT)
  hist(res$total_scores, breaks = "Sturges", col = "#2F4A6B", border = "white",
       xlab = "총점", ylab = "응시자 수", main = main_title)
}

# ------------------------------------------------------------------------------
# HTML 보고서 생성
# ------------------------------------------------------------------------------
esc <- function(x){
  x <- as.character(x)
  x <- gsub("&", "&amp;", x, fixed = TRUE)
  x <- gsub("<", "&lt;", x, fixed = TRUE)
  x <- gsub(">", "&gt;", x, fixed = TRUE)
  x
}

html_table_plain <- function(df){
  hdr <- paste0("<tr>", paste0("<th>", esc(colnames(df)), "</th>", collapse = ""), "</tr>")
  rows <- apply(df, 1, function(r) paste0("<tr>", paste0("<td>", esc(r), "</td>", collapse = ""), "</tr>"))
  paste0("<table><thead>", hdr, "</thead><tbody>", paste(rows, collapse = ""), "</tbody></table>")
}

html_table_distractor <- function(df){
  hdr <- paste0("<tr>", paste0("<th>", esc(colnames(df)), "</th>", collapse = ""), "</tr>")
  cn <- colnames(df)
  rows <- apply(df, 1, function(r){
    correct <- r[["정답"]]
    flagged <- if(!is.na(r[["역작동의심"]]) && r[["역작동의심"]] != "-") strsplit(r[["역작동의심"]], ",")[[1]] else character(0)
    cells <- vapply(cn, function(colname){
      val <- r[[colname]]
      cls <- ""
      if(grepl("^선지\\d+\\(%\\)$", colname)){
        opt_num <- sub("^선지(\\d+)\\(%\\)$", "\\1", colname)
        if(!is.na(correct) && opt_num == correct) cls <- "correct"
        else if(opt_num %in% flagged) cls <- "flag"
      }
      sprintf('<td class="%s">%s</td>', cls, esc(val))
    }, character(1))
    paste0("<tr>", paste(cells, collapse = ""), "</tr>")
  })
  paste0("<table><thead>", hdr, "</thead><tbody>", paste(rows, collapse = ""), "</tbody></table>")
}

png_to_data_uri <- function(plot_fun, width = 900, height = 560){
  tmp <- tempfile(fileext = ".png")
  grDevices::png(tmp, width = width, height = height, res = 130)
  plot_fun()
  grDevices::dev.off()
  raw <- readBin(tmp, "raw", n = file.info(tmp)$size)
  uri <- paste0("data:image/png;base64,", base64enc::base64encode(raw))
  unlink(tmp)
  uri
}

# 영역이 2개 이상이면 표를 영역별 소제목 + 표로 나누고("영역" 열은 중복이라 제거),
# 두 번째 영역부터는 새 페이지로 넘겨 영역별로 따로 보기 좋게 합니다.
# 첫 영역도 h3+표를 한 덩어리로 묶어(page-break-inside:avoid) 남은 공간에 다 들어가지
# 않으면 표 중간이 아니라 통째로 다음 페이지로 넘어가도록 합니다.
# 영역이 1개뿐이면 원래대로 표 하나만 반환합니다.
build_domain_sections <- function(df, domain_order, table_fn){
  domains_present <- domain_order[domain_order %in% unique(df$영역)]
  if(length(domains_present) <= 1) return(table_fn(df))
  parts <- lapply(seq_along(domains_present), function(i){
    dm <- domains_present[i]
    sub <- df[df$영역 == dm, , drop = FALSE]
    sub2 <- sub[, setdiff(colnames(sub), "영역"), drop = FALSE]
    block_cls <- if(i == 1) "domain-block" else "domain-block domain-block-break"
    sprintf('<div class="%s"><h3 class="domain-h">%s <span class="domain-n">(%d문항)</span></h3>%s</div>',
            block_cls, esc(dm), nrow(sub2), table_fn(sub2))
  })
  paste(unlist(parts), collapse = "")
}

HTML_CSS <- "
:root{
  --paper:#F4F5F1; --ink:#1B1F1D; --ink-soft:#54605A;
  --line:#DBDED6; --steel-deep:#213650; --gold:#B9862F;
  --good-soft:#E1EFE6; --warn-soft:#F5E9D3; --bad-soft:#F5E1DC;
}
*{box-sizing:border-box;}
body{font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;color:var(--ink);margin:0;background:#fff;font-size:14px;}
/* A4 인쇄 가능 폭(210mm - 좌우 여백 14mm*2 ≈ 687px)보다 여유 있게 좁혀 잘림 방지 */
.wrap{max-width:650px;margin:0 auto;padding:0 26px 40px;}
header.cover{background:var(--steel-deep);color:#fff;padding:46px 30px;text-align:center;}
header.cover .eyebrow{font-size:12.5px;letter-spacing:.14em;color:#B9C6D6;text-transform:uppercase;}
header.cover h1{font-size:27px;margin:10px 0 6px;}
.cover-sub{font-size:13px;color:#B9C6D6;margin-top:-2px;}
header.cover .meta{font-size:13.5px;color:#C9D2DC;margin-top:12px;}
.gold-bar{height:5px;background:var(--gold);}
section{margin-top:32px;}
section.new-page{page-break-before:always;break-before:page;padding-top:20px;}
.sec-head{display:flex;align-items:baseline;gap:10px;border-bottom:2px solid var(--ink);padding-bottom:7px;margin-bottom:14px;}
.sec-num{color:var(--gold);font-weight:700;font-size:13.5px;}
.sec-head h2{font-size:19px;margin:0;}
.sec-sub{margin-left:auto;font-size:13px;color:var(--ink-soft);}
.stat-cards{display:flex;flex-wrap:wrap;gap:10px;}
.stat-card{flex:1 1 30%;min-width:150px;box-sizing:border-box;border:1px solid var(--line);padding:12px 8px;text-align:center;border-radius:3px;page-break-inside:avoid;break-inside:avoid;}
.stat-card .label{font-size:12.5px;color:var(--ink-soft);}
.stat-card .value{font-size:22px;font-weight:700;color:var(--steel-deep);margin-top:4px;}
.alpha-bar{margin-top:18px;height:14px;border-radius:8px;overflow:hidden;border:1px solid var(--line);position:relative;
  background:linear-gradient(to right, var(--bad-soft) 0%, var(--bad-soft) 60%, var(--warn-soft) 60%, var(--warn-soft) 70%,
  var(--good-soft) 70%, var(--good-soft) 80%, #CFE4D8 80%, #CFE4D8 100%);}
.alpha-marker{position:absolute;top:-4px;width:2px;height:22px;background:var(--steel-deep);}
.domain-block{page-break-inside:avoid;break-inside:avoid;}
.domain-block.domain-block-break{page-break-before:always;break-before:page;}
h3.domain-h{margin:22px 0 8px;font-size:15px;font-weight:700;color:var(--steel-deep);border-left:4px solid var(--gold);padding-left:9px;}
h3.domain-h .domain-n{font-weight:400;color:var(--ink-soft);font-size:12.5px;margin-left:6px;}
table{width:100%;border-collapse:collapse;font-size:13px;margin-top:8px;table-layout:auto;}
th{background:var(--steel-deep);color:#fff;padding:6px 6px;text-align:left;font-size:12px;white-space:nowrap;}
td{padding:5px 6px;border-bottom:1px solid var(--line);}
tr{page-break-inside:avoid;break-inside:avoid;}
tr:nth-child(even) td{background:#F7F7F5;}
td.correct{background:var(--good-soft) !important;font-weight:700;}
td.flag{background:var(--bad-soft) !important;font-weight:700;}
.caption{font-size:13px;color:var(--ink-soft);margin-top:9px;line-height:1.6;}
.callout{background:#FBFBF9;border:1px solid var(--line);border-radius:3px;padding:14px 16px;margin-top:14px;font-size:13.5px;line-height:1.85;page-break-inside:avoid;break-inside:avoid;}
.callout h3{margin:0 0 8px;font-size:15px;}
img.plot{width:100%;border:1px solid var(--line);border-radius:3px;margin-top:12px;page-break-inside:avoid;break-inside:avoid;}
@page{ size:A4; margin:16mm 14mm; }
@media print{ body{font-size:14px;} }
"

build_html_report <- function(res, report_title = ""){
  af <- suppressWarnings(as.numeric(res$summary_df$값[3]))
  af_fmt  <- sprintf("%.2f", af)
  af_word <- if(is.na(af)) "" else if(af >= 0.80) "양호한" else if(af >= 0.70) "수용 가능한" else if(af >= 0.60) "다소 낮은" else "낮은"
  af_pct  <- if(is.na(af)) 0 else max(0, min(1, af)) * 100

  n_examinees <- res$summary_df$값[1]
  n_items     <- res$summary_df$값[2]
  tmean_txt   <- res$summary_df$값[4]
  tsd_txt     <- res$summary_df$값[5]
  trange_txt  <- res$summary_df$값[6]
  tmean <- suppressWarnings(as.numeric(tmean_txt))
  tsd   <- suppressWarnings(as.numeric(tsd_txt))
  ratio <- suppressWarnings(as.numeric(n_examinees) / as.numeric(n_items))

  scatter_uri <- png_to_data_uri(function() make_scatter(res, main_title = "\uBB38\uD56D \uB09C\uC774\uB3C4-\uBCC0\uBCC4\uB3C4 \uC0B0\uC810\uB3C4"))
  hist_uri    <- png_to_data_uri(function() make_hist(res, main_title = "\uCD1D\uC810 \uBD84\uD3EC"))

  domain_html   <- html_table_plain(res$domain_summary)
  item_html     <- build_domain_sections(res$item_stats, res$domain_summary$영역, html_table_plain)
  criteria_html <- html_table_plain(CRITERIA_TABLE)
  dist_html     <- if(!is.null(res$distractor_df)) build_domain_sections(res$distractor_df, res$domain_summary$영역, html_table_distractor) else "<p>선택지 반응분포를 계산할 수 없습니다.</p>"

  report_title <- trimws(report_title)
  has_custom_title <- nzchar(report_title)
  main_title <- if(has_custom_title) esc(report_title) else "CTT 문항분석 결과보고서"
  sub_title_html <- if(has_custom_title) '<div class="cover-sub">CTT 문항분석 결과보고서</div>' else ""

  cover <- sprintf('
<header class="cover">
  <div class="eyebrow">CLASSICAL TEST THEORY &middot; 고전검사이론 문항분석</div>
  <h1>%s</h1>
  %s
  <div class="meta">생성일 %s &middot; 응시자 %s명 &middot; 문항 %s개 &middot; Cronbach\u2019s &alpha; %s</div>
</header>
<div class="gold-bar"></div>', main_title, sub_title_html, format(Sys.Date(), "%Y-%m-%d"), n_examinees, n_items, af_fmt)

  sec01 <- '
<section>
  <div class="sec-head"><span class="sec-num">01</span><h2>보고서 안내</h2></div>
  <p style="font-size:13.5px;line-height:1.9;">
    이 보고서는 고전검사이론(Classical Test Theory)에 기반해 문항과 검사 전체의 특성을 분석한 결과입니다.<br><br>
    &middot; <b>문항난이도(p)</b> &mdash; 응시자 중 해당 문항을 맞힌 비율. 1에 가까울수록 쉬운 문항.<br>
    &middot; <b>전체상관(변별도)</b> &mdash; 문항 점수가 전체(모든 영역 합산) 총점과 함께 움직이는 정도.<br>
    &middot; <b>영역상관</b> &mdash; 문항 점수가 그 문항이 속한 영역의 총점과 함께 움직이는 정도. 영역별로 성격이
    뚜렷이 다른 시험(예: 과목이 섞인 검사)에서는 전체상관보다 영역상관이 더 적절한 변별도 지표일 수 있습니다.<br>
    &middot; <b>Cronbach\u2019s &alpha;</b> &mdash; 문항들이 하나의 특성을 얼마나 일관되게 측정하는지 나타내는 신뢰도 지표.<br>
    &middot; <b>선택지 반응분포</b> &mdash; 선지별 선택 비율. 상위권이 특정 오답을 하위권보다 많이 고르면 문항&middot;정답 오류 의심.
  </p>
  <p style="font-size:13px;color:var(--ink-soft);">※ 문항분석표에는 자동 판정(등급)을 매기지 않았습니다. 해석 기준은 03 판정기준표를 참고해 주세요.</p>
</section>'

  sec02 <- sprintf('
<section>
  <div class="sec-head"><span class="sec-num">02</span><h2>검사 전체 요약</h2><span class="sec-sub">응시자 %s명 &middot; 문항 %s개</span></div>
  <div class="stat-cards">
    <div class="stat-card"><div class="label">응시자 수</div><div class="value">%s</div></div>
    <div class="stat-card"><div class="label">문항 수</div><div class="value">%s</div></div>
    <div class="stat-card"><div class="label">Cronbach\u2019s &alpha;</div><div class="value">%s</div></div>
    <div class="stat-card"><div class="label">총점 평균</div><div class="value">%s</div></div>
    <div class="stat-card"><div class="label">총점 표준편차</div><div class="value">%s</div></div>
    <div class="stat-card"><div class="label">총점 범위</div><div class="value" style="font-size:14px;">%s</div></div>
  </div>
  <div class="alpha-bar"><div class="alpha-marker" style="left:%.1f%%;"></div></div>
  <div class="caption" style="display:flex;justify-content:space-between;"><span>0.0</span><span>.60</span><span>.70</span><span>.80</span><span>1.0</span></div>
  <div class="callout">
    <h3>주요 지표 해설</h3>
    &middot; 이번 검사의 Cronbach\u2019s &alpha;는 %s로, 일반적 기준(.70/.80)에서 %s 신뢰도 수준입니다.<br><br>
    &middot; 응시자 %s명이 문항 %s개에 응답했으며, 문항 대비 응시자 비율은 약 %.1f : 1 입니다.<br><br>
    &middot; 총점은 평균 %s점, 표준편차 %s점입니다. 정규분포를 가정하면 대략 %d~%d점 구간에 응시자의 약 68%%가 분포하는 것으로 추정할 수 있습니다.
  </div>
</section>',
    n_examinees, n_items,
    n_examinees, n_items, af_fmt, tmean_txt, tsd_txt, trange_txt,
    af_pct,
    af_fmt, af_word, n_examinees, n_items, ratio,
    tmean_txt, tsd_txt, round(tmean - tsd), round(tmean + tsd)
  )

  sec03 <- sprintf('
<section class="new-page">
  <div class="sec-head"><span class="sec-num">03</span><h2>판정기준표 (참고용)</h2></div>
  <p style="font-size:13px;color:var(--ink-soft);">문항별 분석 결과에는 아래 기준을 자동으로 적용하지 않았습니다. 해석 시 참고 자료로만 활용해 주세요.</p>
  %s
</section>', criteria_html)

  sec04 <- sprintf('
<section>
  <div class="sec-head"><span class="sec-num">04</span><h2>영역별 요약</h2></div>
  %s
  <div class="caption">영역신뢰도(alpha)·평균영역상관은 해당 영역의 문항 수가 적을수록(2~3개) 불안정할 수 있어 참고용으로만 활용하세요.</div>
</section>', domain_html)

  sec05 <- sprintf('
<section>
  <div class="sec-head"><span class="sec-num">05</span><h2>문항별 분석 결과</h2></div>
  %s
  <div class="caption">지표 해석 기준은 03 판정기준표(참고용)를 참고해 주세요. (이 표에는 등급을 매기지 않았습니다)<br>
  전체상관은 전체 총점 기준, 영역상관은 해당 문항이 속한 영역의 총점 기준입니다. 영역별 성격이 뚜렷이 다른 시험이라면 영역상관을 더 참고해 주세요.</div>
</section>', item_html)

  sec06 <- sprintf('
<section class="new-page">
  <div class="sec-head"><span class="sec-num">06</span><h2>문항별 선택지(선지) 반응분포</h2></div>
  %s
  <div class="caption">초록 셀 = 정답 선지 &middot; 빨간 셀(역작동의심) = 상위 집단이 하위 집단보다 해당 오답 선지를 더 많이 선택한 경우입니다.</div>
</section>', dist_html)

  sec07 <- sprintf('
<section class="new-page">
  <div class="sec-head"><span class="sec-num">07</span><h2>문항 분포 시각화</h2></div>
  <img class="plot" src="%s">
  <img class="plot" src="%s">
</section>', scatter_uri, hist_uri)

  footer <- '<p style="text-align:center;color:var(--ink-soft);font-size:10.5px;margin-top:30px;">CTT Analyzer &middot; R + CTT package 기반</p>'

  paste0(
    '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><title>', main_title, '</title><style>',
    HTML_CSS, '</style></head><body>',
    cover,
    '<div class="wrap">',
    sec01, sec02, sec03, sec04, sec05, sec06, sec07,
    footer,
    '</div></body></html>'
  )
}

# ------------------------------------------------------------------------------
# UI
# ------------------------------------------------------------------------------
ui <- fluidPage(
  tags$head(tags$style(HTML("
    * { font-family: '맑은 고딕', 'Malgun Gothic', sans-serif; }
    body{background:#F4F5F1;}
    .title-box{background:#213650;color:#fff;padding:20px 26px;border-radius:4px;margin-bottom:20px;}
    .title-box h2{margin:0;} .title-box p{margin:4px 0 0;color:#C9D2DC;font-size:13px;}
    .well{background:#FBFBF9;}
    .btn-primary{background-color:#213650;border-color:#213650;}
  "))),
  div(class = "title-box",
      h2("CTT 문항분석기 — Shiny GUI"),
      p("CTT 패키지 계산 결과를 파일 업로드만으로 확인하고 다운로드하세요.")
  ),
  sidebarLayout(
    sidebarPanel(
      width = 3,
      fileInput("dataFile", "① 응답 데이터 파일 (.xlsx/.csv)", accept = c(".xlsx", ".xls", ".csv")),
      uiOutput("dataSheetUI"),
      fileInput("keyFile", "② 정답표 파일 (.xlsx/.csv, 영역 행/열 포함 가능)", accept = c(".xlsx", ".xls", ".csv")),
      uiOutput("keySheetUI"),
      checkboxInput("hasId", "첫 열이 응시자 ID", value = FALSE),
      selectInput("nGroups", "선택지 분포 분석 집단 수", choices = c(2, 3, 4), selected = 4),
      actionButton("runBtn", "CTT 분석 실행", class = "btn-primary"),
      hr(),
      downloadButton("downloadXlsx", "엑셀 다운로드"),
      br(), br(),
      textInput("reportTitle", "보고서 제목 (선택)", value = "",
                placeholder = "예: 2026년 하반기 신입사원 채용 필기시험"),
      downloadButton("downloadHtml", "HTML 보고서 다운로드"),
      p(style = "font-size:11px;color:#777;margin-top:8px;",
        "HTML 보고서를 브라우저에서 열어 인쇄(Ctrl+P → PDF로 저장)하면 PDF로도 저장할 수 있습니다.")
    ),
    mainPanel(
      width = 9,
      verbatimTextOutput("statusMsg"),
      tabsetPanel(
        tabPanel("요약", tableOutput("summaryTable")),
        tabPanel("영역별 요약", tableOutput("domainTable")),
        tabPanel("문항분석", DTOutput("itemTable")),
        tabPanel("판정기준표(참고용)", tableOutput("criteriaTable")),
        tabPanel("선택지반응분포", DTOutput("distTable")),
        tabPanel("그래프", plotOutput("scatterPlot", height = "380px"),
                            plotOutput("histPlot", height = "380px"))
      )
    )
  )
)

# ------------------------------------------------------------------------------
# Server
# ------------------------------------------------------------------------------
server <- function(input, output, session){

  output$dataSheetUI <- renderUI({
    req(input$dataFile)
    ext <- tolower(tools::file_ext(input$dataFile$name))
    if(ext %in% c("xlsx", "xls")){
      sheets <- excel_sheets(input$dataFile$datapath)
      if(length(sheets) > 1) selectInput("dataSheet", "응답 데이터 시트", choices = sheets, selected = sheets[1])
    }
  })

  output$keySheetUI <- renderUI({
    req(input$keyFile)
    ext <- tolower(tools::file_ext(input$keyFile$name))
    if(ext %in% c("xlsx", "xls")){
      sheets <- excel_sheets(input$keyFile$datapath)
      if(length(sheets) > 1){
        guess <- sheets[grepl("정답|answer|key", sheets, ignore.case = TRUE)]
        sel <- if(length(guess) > 0) guess[1] else sheets[1]
        selectInput("keySheet", "정답표 시트", choices = sheets, selected = sel)
      }
    }
  })

  result <- eventReactive(input$runBtn, {
    req(input$dataFile, input$keyFile)
    ext_d <- tolower(tools::file_ext(input$dataFile$name))
    sheet_d <- if(!is.null(input$dataSheet)) input$dataSheet else NULL
    raw_df <- read_table_any(input$dataFile$datapath, ext_d, sheet_d)

    ext_k <- tolower(tools::file_ext(input$keyFile$name))
    sheet_k <- if(!is.null(input$keySheet)) input$keySheet else NULL
    key_df <- read_table_any(input$keyFile$datapath, ext_k, sheet_k)

    tryCatch({
      run_ctt(raw_df, key_df, input$hasId, as.integer(input$nGroups))
    }, error = function(e){
      showNotification(paste("오류:", conditionMessage(e)), type = "error", duration = NULL)
      NULL
    })
  })

  output$statusMsg <- renderText({
    res <- result()
    if(is.null(res)) return("응답 데이터와 정답표를 업로드하고 [CTT 분석 실행]을 눌러주세요.")
    sprintf("완료 · 응시자 %s명 · 문항 %s개 · Cronbach's alpha = %s",
            res$summary_df$값[1], res$summary_df$값[2], res$summary_df$값[3])
  })

  output$summaryTable  <- renderTable({ req(result()); result()$summary_df })
  output$domainTable   <- renderTable({ req(result()); result()$domain_summary })
  output$criteriaTable <- renderTable({ CRITERIA_TABLE })
  output$itemTable     <- renderDT({ req(result()); datatable(result()$item_stats, options = list(pageLength = 20)) })
  output$distTable     <- renderDT({
    res <- result(); req(res)
    if(is.null(res$distractor_df)) return(datatable(data.frame(안내 = "선택지 반응분포를 계산할 수 없습니다.")))
    datatable(res$distractor_df, options = list(pageLength = 25))
  })

  output$scatterPlot <- renderPlot({ req(result()); make_scatter(result()) })
  output$histPlot     <- renderPlot({ req(result()); make_hist(result()) })

  output$downloadXlsx <- downloadHandler(
    filename = function() "CTT_분석결과.xlsx",
    content = function(file){
      res <- result(); req(res)
      sheets <- list(요약 = res$summary_df, 영역별요약 = res$domain_summary,
                     문항분석 = res$item_stats, 판정기준표 = CRITERIA_TABLE)
      if(!is.null(res$distractor_df)) sheets[["선택지반응분포"]] <- res$distractor_df
      write_xlsx(sheets, file)
    }
  )

  output$downloadHtml <- downloadHandler(
    filename = function(){
      title <- trimws(input$reportTitle)
      if(nzchar(title)){
        safe <- gsub('[\\\\/:*?"<>|]+', "_", title)
        safe <- gsub("\\s+", "_", safe)
        paste0(safe, "_CTT_분석결과.html")
      } else {
        "CTT_분석결과.html"
      }
    },
    content = function(file){
      res <- result(); req(res)
      writeLines(build_html_report(res, report_title = input$reportTitle), file, useBytes = TRUE)
    }
  )
}

shinyApp(ui, server)
