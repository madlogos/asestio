#' Simplify Checkup Data Sumamry
#'
#' Simplify the summary column of checkup raw dataset.
#' @details There are 3 additional arguments: \cr
#' \describe{
#'   \item{beginner}{The beginner RegEx expression of a sentence.}
#'   \item{ender}{The ending RegEx expression of a sentence.}
#'   \item{special.chr}{Invoke the beginner/ender structure and keep the special.chr.}
#' }
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param summ.col The index or name of the summary column
#' @param file The checkup raw dataset in Excel spreadsheet format
#' @param sheet The index or name of the sheet. Default 1.
#' @param provider A string. Currently support "Ciming", "Meinian" and "Ikang"
#'
#' @return A xlsx file under the same directory, containing modified sumamry column
#' @export
#' @importFrom compiler cmpfun
#' @importFrom readxl read_excel
#' @importFrom readr write_csv write_excel_csv
#' @importFrom openxlsx write.xlsx
#' @importFrom stringi stri_enc_toutf8
#' @importFrom cellranger letter_to_num
#'
#' @examples
#' \dontrun{
#' simpSum("C", provider="Ikang")
#'
#' ##
#' }
simpSumm <- function(
    summ.col, file=NULL, sheet=1, provider=c("Ciming", "Meinian", "Ikang"),
...) {
    provider <- match.arg(provider)
    if (is.null(file)) file <- file.choose()
    stopifnot(grepl("\\.(csv|xls|xlsx)$", file))
    dat <- if (grep("\\.csv$", file)) read_csv(file) else
        read_excel(file, sheet, col_names=FALSE)
    summ.col <- letter_to_num(summ.col)
    orig.txt <-  dat[, summ.col]

    txt <-  unlist(orig.txt)
    txt <-  stri_enc_toutf8(txt)
    class(txt) <- provider

    out <- simpStr(txt, ...)
    write.xlsx(data.frame(summ=unlist(out), stringsAsFactors=FALSE),
               file=paste0(sub("^(.+)\\.(csv|xls|xlsx)", "\\1", file),
                           "_Summ.xlsx"))
}

#' @export
#' @rdname simpSumm
simp_summ <- simpSumm

#' @importFrom stringr str_replace_all str_split str_count str_replace_na str_replace
#' @importFrom stringr str_detect str_extract str_trim
simpStr <- function(str, ...) {
	UseMethod("simpStr")
}

simpStr.Ciming <- function(
    str, beginner="\\d+\\.\\D|\\(\\d+\\)\\D", ender="[\u{3002}\u{ff1f}]",
    special.chr="\u77eb\u6b63\u89c6\u529b\u4e0d\u8db3|\u60a8\u7684\u4f53\u91cd.+?\u8303\u56f4"
){
    ## mark clause beginners
    staging <- str_replace_all(str, paste0("\\s*(", beginner, ")"), "\\\n\\1")
    staging <- str_replace_all(staging, "\\\r", "\\\n")
    staging <- str_replace_all(staging, "(\\\n){2,}", "\\\n")
    ## split the string by beginners
    staging <- str_split(staging, "\\r|\\n")
    ## simplify the wording
    staging <- lapply(staging, function(chr){
        has.beginner = str_count(chr, beginner) > 0
        has.ender = str_count(chr, ender) > 0
        if (is.null(special.chr) || is.na(special.chr) || special.chr=="") {
            chr[has.ender] = str_extract(
                chr[has.ender], paste0("^[^", ender, "]+?", ender))
        }else{
            has.spec = str_count(chr, special.chr) > 0
            chr[has.spec] = str_extract(
                chr[has.spec], paste0("^.+?(", special.chr, ")"))
            chr[has.ender & !has.spec] = str_extract(
                chr[has.ender & !has.spec], paste0("^[^", ender, "]+?", ender))
        }
        return(chr)
    })
    ## combine results
    output <- lapply(staging, function(s) str_replace_na(unlist(s), ""))
    output <- unlist(lapply(output, str_c, collapse="\\r\\n"))

    ## fortify
    output <- str_replace_all(output, "\\d\\.\u5bb6\u65cf\u53f2\u63d0\u793a.+$", "")
    output <- str_replace_all(output, "^\\s+|\\s*(\\d+\\.)\\s*|^\\\\r\\\\n|\\\\r$|\\\\n$", "\\1")
    return(output)
}
simpStr.Ikang <- function(
    str, beginner="\u3010", ender="[\uff1a:]", special.chr=NULL
){
    ## pre-process
    staging <- str_replace_all(str, "\u3011\\r\\n", "\u3011")

    ## mark clause beginners
    staging <- str_replace_all(str, paste0("\\s*(", beginner, ")"), "\\\n\\1")
    staging <- str_replace_all(staging, "\\\r", "\\\n")
    staging <- str_replace_all(staging, "(\\\n){2,}", "\\\n")
    ## split the string by beginners
    staging <- str_split(staging, "\\r|\\n")
    ## simplify the wording
    staging <- lapply(staging, function(chr){
        has.beginner = str_count(chr, beginner) > 0
        has.ender = str_count(chr, ender) > 0
        if (is.null(special.chr) || is.na(special.chr) || special.chr=="") {
            chr[has.ender] = str_extract(
                chr[has.ender], paste0("^[^", ender, "]+?", ender))
        }else{
            has.spec = str_count(chr, special.chr) > 0
            chr[has.spec] = str_extract(
                chr[has.spec], paste0("^.+?(", special.chr, ")"))
            chr[has.ender & !has.spec] = str_extract(
                chr[has.ender & !has.spec], paste0("^[^", ender, "]+?", ender))
        }
        chr = chr[has.beginner]
        return(chr)
    })
    ## combine results
    output <- lapply(staging, function(s) str_replace_na(unlist(s), ""))
    output <- unlist(lapply(output, str_c, collapse="\\r\\n"))

    ## fortify

    output <- str_replace_all(output, "(\\\\r\\\\n){2,}", "\\\r\\\n")
    output <- str_replace_all(
        output, "^\\s+|\\s*(\\d+\\.)\\s*|^\\\\r\\\\n|\\\\r$|\\\\n$|\\\\r\\\\n$", "\\1")
    return(output)
}
simpStr.Meinian <- function(str){
    ## pre-process
    staging <- str_replace_all(
        str, regex("\u5c0a\u656c\u7684.+?(\u{2605})|\u4f53\u68c0\u7ed3\u679c\u8bf4\u660e.+$",
                   dotall=TRUE), "\\1")

    ## split the string by beginners
    staging <- str_split(staging, "\\r|\\n")
    ## simplify the wording
    staging <- lapply(staging, function(s){
        s <- str_trim(s)
        s[str_detect(s, "^\\d+\u3001.+|^[\uff08\\()]|^\u8bf7|^\u5982|^\u63d0\u793a")] <- ""
        has.futile <- str_detect(
            s, paste0("\u5efa\u8bae.+$|\u8bf7.+$|\u6ce8\u610f.+$|",
                      "\u7532\u72b6\u817a\u56ca\u80bf\u662f.+$|",
                      "\u6162\u6027\u9f3b\u708e\u662f.+$|\u80be\u7ed3\u6676\u591a\u4e3a.+$|",
                      "\u74e3\u819c\u8fd4\u6d41\u591a\u89c1.+$|\u75db\u98ce\u4f9d\u75c5\u56e0.+$"))
        s[has.futile] <- str_replace(
            s[has.futile],
            paste0("\u5efa\u8bae.+$|\u8bf7.+$|\u6ce8\u610f.+$|",
                   "\u7532\u72b6\u817a\u56ca\u80bf\u662f.+$|",
                   "\u6162\u6027\u9f3b\u708e\u662f.+$|\u80be\u7ed3\u6676\u591a\u4e3a.+$|",
                   "\u74e3\u819c\u8fd4\u6d41\u591a\u89c1.+$|\u75db\u98ce\u4f9d\u75c5\u56e0.+$"),
            "")
        s <- s[s != ""]
    })
    ## combine results
    output <- lapply(staging, function(s) str_replace_na(unlist(s), ""))
    output <- unlist(lapply(output, str_c, collapse="\\r\\n"))

    ## fortify
    output <- str_replace_all(output, "(\\\\r\\\\n){2,}", "\\\r\\\n")
    output <- str_replace_all(
        output, "^\\s+|\\s*(\\d+\\.)\\s*|^\\\\r\\\\n|\\\\r$|\\\\n$|\\\\r\\\\n$", "\\1")
    return(output)
}

