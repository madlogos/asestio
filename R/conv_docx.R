#' Convert doc to docx
#'
#' Convert doc to docx in batch manner
#'
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param files character vector of dic files. Default NULL (A GUI wizard is
#' then called).
#' @param delete.original Logical. If TRUE, then the original doc files will be
#' deleted after the conversion. Default FALSE.
#' @param use.parallel Logical, whether or not to use parallel mode when 
#' there are more than 50 files. Default TRUE.
#'
#' @return Nothing
#' @export

#' @import RDCOMClient
#' @importFrom gWidgets2 gfile
#' @importFrom parallel makeCluster parLapplyLB stopCluster
#' @importFrom stringi stri_conv stri_enc_toascii
#' @importFrom stringr str_replace
#'
#' @examples
#' \dontrun{
#' convDoc2Docx()
#' }
convDoc2Docx <- function(files=NULL, delete.original=FALSE, use.parallel=TRUE){
    Sys.setlocale("LC_CTYPE", "Chs")
    # -------ensure Rtools is included in Sys.Path-------
    #addRtoolsPath()
    if (is.null(files)) {
        if (Sys.info()['sysname'] == "Windows"){
            files <- invisible(choose.files(
                paste0(getOption("init.dir"), "*.*"),
                caption="Select .doc files...",
                filters=rbind(matrix(c("doc files(.doc)", "*.doc;*.DOC"), nrow=1),
                              Filters["All",])))
        }else{
            files <- gfile(text="Select .doc files...", type="open",
                           initial.dir=getOption("init.dir"),
                           filter=c('doc files'='doc'),
                           multi=TRUE, toolkit=guiToolkit("tcltk"))
            files <- stri_conv(files, "CP936", "UTF-8")
        }
    }else{
        files <- unlist(files)
        files <- files[grepl("\\.doc$", tolower(files))]
    }
    
    if (any(grepl("[[:cntrl:]]", stri_enc_toascii(files))))
        warning("There are non-ASCII (e.g. Chinese characters) characters ",
                "in the filenames. ", 
                "Please rectify it before continue processing.")
    
    if (any(file.exists(files))){
        wdApp <- COMCreate("Word.Application")
        convDoc <- cmpfun(function(file){
            doc <- wdApp[["Documents"]]$Open(file)
            doc$SaveAs(str_replace(file, "doc$", "docx"), 12)  # wdFormatXMLDocument = 12
            doc$Close()
        })
        
        if (length(files) > 50 && use.parallel){
            ## parallel computation
            cl <- makeCluster(detectCores(
                logical=! Sys.info()[["sysname"]] == "Windows"))
            # clusterExport(cl, "convDoc", envir=environment())
            created <- parLapplyLB(cl, files, convDoc)
            # stopImplicitCluster()
            stopCluster(cl)
        }else{
            created <- lapply(files, convDoc)
        }
        
        wdApp$Quit()
        wdApp <- NULL
        if (delete.original) unlink(files)
        invisible(paste("A total of", seq_len(length(files)), "files conversed!",
                        ifelse(delete.original, "\nOriginal xls files are deleted.",
                               "")))
    }else{
        invisible("No file assigned.")
    }
}

#' @export
#' @rdname convDoc2Docx
doc2docx <- convDoc2Docx