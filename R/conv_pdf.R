#' Convert PDF to Text Files
#'
#' Use Xpdf to convert PDFs to text files. The function relies on
#' \pkg{\link{doParallel}} to convert the PDFs in parallel.
#' It calls Xpdf to complete the conversion. Currently this function
#' only supports Windows.
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param mode integer, Default 1L. \describe{
#' \item{1L}{as is (-raw)} \item{2L}{with layout (-layout)}
#' \item{3L}{no breaks (-nopgbrk)} \item{4L}{without format ()}
#' \item{5L}{as table (-table)} \item{6L}{simplified (-simple)}
#' }
#' @param silent logical, whether show the results of the conversion.
#' Default \code{FALSE}.
#' @param converter which executable program to use to extract the pdf, "pdftotext" 
#' or "pdftohtml", default "pdftotext".
#' @param converter.path Path to the pdftotext.exe or pdftohtml.exe executable
#' program. If NULL, make sure the xpdf folder is under R.home() or the parent 
#' directory of R.home(). E.g., for an x86 Win PC, \code{paste0(R.home(), 
#' paste0("/xpdfbin-win-4.00/bin32/pdftotext.exe"))}.
#' @param use.parallel Logical, whether to apply parallel computation when there 
#' are more than 50 files to process. Default \code{TRUE}.
#'
#' @return Full name of the PDFs if you set \code{silent=FALSE}.
#' @export
#' @aliases conv_pdfs
#' 
#' @importFrom doParallel registerDoParallel stopImplicitCluster
#' @importFrom foreach foreach %dopar%
#' @importFrom parallel makeCluster stopCluster
#' @importFrom gWidgets2 gfile
#' @importFrom compiler cmpfun
#' @importFrom stringi stri_conv
#' @importFrom stringr str_replace str_replace_all
#'
#' @examples
#' \dontrun{
#' convPDFs()
#' }
convPDFs <- function(mode=1L, silent=FALSE, converter=c("pdftotext", "pdftohtml"), 
                     converter.path=NULL, use.parallel=TRUE){
    converter <- match.arg(converter)
    if(Sys.info()[['sysname']] != "Windows")
        stop("This function only supports Windows platform currently.")
    stopifnot(mode %in% 1L:6L)
    modes <- c("   1: as is (-raw)", "   2: with format (-layout)", 
               "   3: no page breaks (-nopgbrk)",
               "   4: without format ", "   5: as table (-table)", 
               "   6: simplified (-simple)")
    modes[mode] <- sub("^ {2}", "->", modes[mode])
    if (converter == "pdftotext"){
        invisible(message("PDF transformation mode:\n",
                          paste(modes, collapse="\n")))
        if (missing(mode)){
            mode.input <- as.integer(readline(
                "Input the conversion mode (1-6), or type enter to bypass:\n"))
        }
        if (!is.na(mode.input)){
            mode <- mode.input %>% gsub("->", "", .)
            modes[mode] <- sub("^ {2}", "->", modes[mode])
            invisible(message("PDF transformation mode:\n",
                              paste(modes, collapse="\n")))
        }
    }
    
    #Sys.setlocale("LC_ALL", "Chs")  # deprecated
    if (is.null(converter.path)) 
        converter.path <- paste0(searchXpdfFolder(), "/", converter, ".exe")
    if (! file.exists(converter.path))
        stop(paste(
            "The converter program is not found! ",
            "Please install xpdfbin-win-4.00 to the R root path or its parent directory.",
            "Type R.home() to find more details.",
            sep="\n"))
    
    #---------get pdfs------------
    if (Sys.info()['sysname'] == "Windows"){
        pdfs <- invisible(choose.files(
            paste0(getOption("init.dir"), "*.*"), caption="Select PDF files...",
            filters=Filters[c("pdf", "All"),]))
    }else{
        pdfs <- gfile(text = "Select PDF files...", type = "open",
                      initial.dir = getOption("init.dir"),
                      filter = c("PDF Files"="pdf"), 
                      multi=TRUE, toolkit=guiToolkit(getOption("guiToolkit")))
        pdfs <- stri_conv(pdfs, "CP936", "UTF-8")
    }
    
    #-------------Conversion-----------
    # pdfs<-list.files(filepath,pattern="\\.pdf$",recursive=TRUE,full.names=TRUE)
    convPDF <- cmpfun(invisible(convPDF))
    if (any(file.exists(pdfs))){
        pdfs <- str_replace_all(pdfs, "\\\\", "/")
        pdf.path <- str_replace(pdfs[[1]], "^(.+/)[^/]+$", "\\1")
        
        pdfs <- pdfs[file.exists(pdfs)]
        if (length(pdfs) > 0){
            if (length(pdfs) > 50 && use.parallel){
                ## cluster computation
                cl <- makeCluster(detectCores(
                    logical=! Sys.info()[["sysname"]] == "Windows"))
                registerDoParallel(cl, detectCores(
                    logical=! Sys.info()[["sysname"]] == "Windows"))
                invisible(
                    foreach(i = 1:length(pdfs), .export=c(
                        "mode", "converter", "converter.path")) %dopar% 
                        convPDF(pdfs[i], mode=mode, converter=converter, 
                                converter.path=converter.path))
                stopImplicitCluster()
                stopCluster(cl)
            }else{
                invisible(lapply(pdfs, convPDF, mode=mode, converter=converter,
                                 converter.path=converter.path))
            }
        }
        invisible(message(paste("Done. A total of", length(pdfs),
                                "pdf files conversed in", pdf.path, ".")))
        if (! silent) return(pdfs) else return(NULL)
    }else{
        invisible(message("No file assigned!"))
        return(NULL)
    }
}

#' @export
#' @rdname convPDFs
conv_pdfs <- convPDFs

convPDF <- function(pdf, mode=1, converter=c("pdftotext", "pdftohtml"),
                    converter.path=NULL){
    converter <- match.arg(converter)
    if (converter == "pdftotext"){
        switch(mode,
               #------1 = as is------
               system(paste0("\"", converter.path, "\" -raw \"", pdf, "\""),
                      wait = FALSE),
               #------2 = with format-----
               system(paste0("\"", converter.path, "\" -layout \"", pdf, "\""),
                      wait = FALSE),
               #------3 = no breaks------
               system(paste0("\"", converter.path, "\" -nopgbrk \"", pdf, "\""),
                      wait = FALSE),
               #------4 = without format
               system(paste0("\"", converter.path, "\" \"", pdf, "\""),
                      wait = FALSE),
               #------5 = as table------
               system(paste0("\"", converter.path, "\" -table \"", pdf, "\""),
                      wait = FALSE),
               #------6 = simplified------
               system(paste0("\"", converter.path, "\" -simple \"", pdf, "\""),
                      wait = FALSE)
        )
    } else {
        system(paste0("\"", converter.path, "\" \"", pdf, "\" \"", 
                      sub("^(.+)\\.pdf$", "\\1", pdf),  "\""), 
               wait = FALSE)
    }
}
