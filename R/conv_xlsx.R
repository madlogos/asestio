#' Convert xls to xlsx
#'
#' Convert xls to xlsx in batch manner
#'
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param files character vector of xls files. Default NULL (A GUI wizard is
#' then called).
#' @param delete.original Logical. If TRUE, then the original xls files will be
#' deleted after the conversion.
#' @param use.parallel Logical, if use parallel mode when there are more than
#' 50 files. Default TRUE.
#'
#' @return Nothing
#' @export
#' @importFrom gWidgets2 gfile
#' @importFrom parallel makeCluster mcmapply stopCluster
#' @importFrom stringi stri_conv
#' @importFrom stringr str_replace_all
#' @importFrom rio convert
#'
#' @examples
#' \dontrun{
#' convXls2Xlsx()
#' }
convXls2Xlsx <- function(files=NULL, delete.original=FALSE, use.parallel=TRUE){  # delete xls conversed?
    Sys.setlocale("LC_CTYPE", "Chs")
    # -------ensure Rtools is included in Sys.Path-------
    
    if (is.null(files)) {
        if (Sys.info()['sysname'] == "Windows"){
            files <- invisible(choose.files(
                "//ship-oa-001/China_Health_Advisory/Analytics/*.*",
                caption="Select .xls files...",
                filters=rbind(matrix(c("xls files(.xls)", "*.xls;*.XLS"), nrow=1),
                              Filters["All",])))
        }else{
            files <- gfile(text="Select .xls files...", type="open",
                       initial.dir="//ship-oa-001/China_Health_Advisory/Analytics/",
                       filter=c('xls files'='xls'),
                       multi=TRUE, toolkit=guiToolkit("tcltk"))
            files <- stri_conv(files, "CP936", "UTF-8")
        }
    }else{
        files <- tolower(unlist(files))
        files <- files[grepl("\\.xls$", files)]
    }
    
    if (any(file.exists(files))){
        if (length(files) > 50 && use.parallel){
            ## parallel computation
            cl <- makeCluster(detectCores(
                logical=! Sys.info()[["sysname"]] == "Windows"))
            # clusterExport(cl, "rio::convert", envir=environment())
            created <- mcmapply(convert, files,
                                str_replace_all(files, "xls$", "xlsx"))
            # stopImplicitCluster()
            stopCluster(cl) 
        }else{
            created <- mapply(convert, files,
                              str_replace_all(files, "xls$", "xlsx"))
        }
        if (delete.original) unlink(files)
        invisible(paste("A total of", seq_len(length(files)), "files conversed!",
                     ifelse(delete.original, "\nOriginal xls files are deleted.",
                            "")))
    }else{
        invisible("No file assigned.")
    }
}

#' @export
#' @rdname convXls2Xlsx
xls2xlsx <- convXls2Xlsx