#' Convert PDFs to TXTs And Combine Them to A Data.frame
#'
#' This is a function combining \code{\link{convPDFs}} and \code{\link{bindTXTs}}.
#' You can directly select PDFs and then convert them to txts, based on which you
#' can then combine them to a data.frame with a mapping rule.file.
#'
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param mode integer, Default 1L. \describe{
#' \item{1L}{as is (-raw)} \item{2L}{with layout (-layout)}
#' \item{3L}{no breaks (-nopgbrk)} \item{4L}{without format ()}
#' \item{5L}{as table (-table)} \item{6L}{simplified (-simple)}
#' }
#' @param converter The converter program to use, either "pdftotext" or "pdftohtml"
#' @param converter.path Path to the converter program. If NULL, make sure the
#' xpdf folder is under R.home(). E.g., \code{paste0(R.home(),
#' paste0("/xpdfbin-win-4.00/bin64/pdftotext.exe"))}
#' @param rule.file Full path of the conversion dictionary file. Default
#' \code{paste(path, <dict name>)}.
#' @param path Path of the directory to the dictionary file.
#'
#' @return Nothing
#' @export
#'
#' @examples
#' \dontrun{
#' bindPDFs()
#' }
bindPDFs <- function(mode=1, converter=c("pdftotext", "pdftohtml"), 
                     converter.path=NULL, 
                     rule.file=paste0(path, 'IndRptToTable(RegExp).xlsx'),
                     path=getOption("toolkit.dir")){
    stopifnot(mode %in% 1:6)
    converter <- match.arg(converter)
    pdfs <- convPDFs(mode, converter=converter, converter.path=converter.path)
    if (!is.null(pdfs)){
        txts <- str_replace_all(pdfs, "^(.+)\\.[Pp][Dd][Ff]$", "\\1\\.txt")
        bindTXTs(rule.file, path, txts=txts)
    }else{
        warning("No PDF is designated!")
    }
}

#' @export
#' @rdname bindPDFs
bind_pdfs <- bindPDFs

