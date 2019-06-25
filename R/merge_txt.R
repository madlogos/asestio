#' Combine Txts And Convert Them to A Data Frame
#'
#' Combine Txts derived from PDFs and convert them to a data frame using a
#' conversion dictionary stored in IndRptToTable(RegExp).xlsx.
#' 
#' The function yields two files: an .xlsx file containing merge data and a .csv
#' containing the unmatched line headers that are potentially new variables.
#'
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param rule.file Full path of the conversion dictionary file. Default
#' \code{paste(path, dict name)}.
#' @param path Path of the directory to the dictionary file.
#' @param txts Txts assigned for combination. If NULL, a wizard will guide you
#' to assign them.
#' @param use.parallel Logical, whether to apply parallel computation when there 
#' are more than 50 files to process. Default \code{TRUE}.
#' 
#' @return Nothing
#' @export
#'
#' @importFrom gWidgets2 gbasicdialog gvbox ggroup gradio gframe gfile visible
#' @importFrom parallel makeCluster clusterExport parLapplyLB parSapply stopCluster
#' @importFrom stringr str_replace str_replace_all str_detect
#' @importFrom stringi stri_conv
#' @importFrom compiler cmpfun
#' @importFrom openxlsx read.xlsx getSheetNames write.xlsx createStyle
#' @importFrom readr read_file
#' @examples
#' \dontrun{
#' bindTXTs()
#' }
bindTXTs <- function(rule.file=paste0(path, "IndRptToTable(RegExp).xlsx"),
                    path=getOption("toolkit.dir"), txts=NULL, use.parallel=TRUE) {

    #1. ----------------Modify locale, load necessary packages------------------
    Sys.setlocale("LC_ALL", "Chinese")
    time.begin <- Sys.time()
    
    # options(guiToolkit="RGtk2")
    if (! file.exists(rule.file))
        stop("The rule.file cannot be found!")
    # addRtoolsPath()  # Add Rtools to Sys.path for openxlsx
    #2. -------------------Read RegExp rules----------------------------------
    sources <- getSheetNames(rule.file)

    ### -----------Open a gui window------------
    .tmpEnv <- new.env()
    .tmpEnv$sources <- NULL
    
    window <- gbasicdialog("Select data source", do.buttons=TRUE, 
                           handler=function(h, ...){
        .tmpEnv$sources <<- enc2native(chkmap$get_value())
    })
    size(window) <- c(200, 200)
    invisible(window)
    #invisible(window$set_size(200, 200))
    box <- gvbox(cont=window)
    # addHandlerChanged(window, handler=function(...){
    #     gtkMainQuit()
    # })
    gg1 <- ggroup(cont=box)
    gg2 <- ggroup(cont=box, horizontal = TRUE)
    box1 <- gvbox(cont=gg1)
    frm1 <- gframe("Sources:", cont=box1)
    chkmap <- gradio(items=sources, selected=1, index=TRUE, cont=frm1)
    box21 <- gvbox(cont=gg2)
    box22 <- gvbox(cont=gg2)
    visible(window)
    
    ### -----------------Close window---------------
    if (is.null(.tmpEnv$sources)) {
        stop("No source designated!")
    }else {
        datasource <- .tmpEnv$sources
        if (! datasource %in% sources)
            datasource <- iconv(.tmpEnv$sources, from="UTF8", to="CP936")

        ##Choose sheet
        rule.st <- read.xlsx(rule.file, sheet=datasource, detectDates=TRUE)

        #3. -----------------Clean txts in loop and cover them------------------
        ## You can only read but skip the clean part
        # txts<-list.files(path,pattern="\\.txt$",recursive=T,full.names=T)
        if (is.null(txts))
            if (Sys.info()['sysname'] == "Windows"){
                txts <- invisible(choose.files(
                    paste0(getOption("init.dir"), "*.*"), 
                    caption="Select .txt files...",
                    filters=Filters[c('txt', "All"),]))
            }else{
                txts <- gfile("Select .txt files...", type='open',
                              initial.dir=getOption("init.dir"),
                              filter=c('Txt files'='txt'),
                              multi=TRUE, toolkit=guiToolkit("tcltk"))
                txts <- stri_conv(txts, "CP936", "UTF-8")
            }

        #4. -----------Go thru docs and extract info to make data.frame--------
        
        if (length(txts) > 0){
            mapTxt2Df <- cmpfun(mapTxt2Df)
            getTxtHead <- cmpfun(getTxtHead)
            if (length(txts) > 50 && use.parallel){
                ## parallel computation
                cl <- makeCluster(detectCores(
                    logical=! Sys.info()[["sysname"]] == "Windows"))
                # registerDoParallel(cl, detectCores(
                #      logical=! Sys.info()[["sysname"]] == "Windows"), nocompile=FALSE)
                # lstconv <- invisible(
                #      foreach(i = 1:length(txts), .export=c("rule.st", "mapTxt2Df"),
                #              .packages="aseshms") %dopar% 
                #          cmpfun(mapTxt2Df(txt=txts[i], ruleDf=rule.st)))
                clusterExport(cl, "rule.st", envir=environment())
                lstconv <- parLapplyLB(cl, txts, mapTxt2Df, ruleDf=rule.st)
                lstunmatch <- parSapply(cl, txts, getTxtHead, ruleDf=rule.st)
                # stopImplicitCluster()
                stopCluster(cl) 
            }else{
                lstconv <- lapply(txts, mapTxt2Df, ruleDf=rule.st)
                lstunmatch <- sapply(txts, getTxtHead, ruleDf=rule.st)
            }
            dfconv <- do.call("rbind", c(lstconv, deparse.level=0))
            unmatched <- unique(unlist(lstunmatch))
            #dfconv[] <- sapply(dfconv, function(column) gsub("\\f", "\\n", column))
        }else{
            dfconv <- NULL
            unmatched <- NULL
        }

        #5. -------------------Export extract to xlsx------------------------
        filepath <- str_replace_all(txts[1], "\\\\", "/")
        filepath <- str_replace_all(filepath, "^(.+/)[^/]+\\.[Tt][Xx][Tt]$", "\\1")

        if (! str_detect(filepath, ".+/$")) filepath <- paste0(filepath,"/")
        if (!is.null(dfconv)) write.xlsx(
            as.data.frame(dfconv), paste0(filepath,"bind.xlsx"), 
            sheetName="Sheet1", showNA=FALSE)
        if (!is.null(dfconv) && file.exists(paste0(filepath, "bind.xlsx"))) {
            message(paste("Combined data set generated:",
                            paste0(filepath, "bind.xlsx"), "!"))
        }else{
            warning("Combined data set failed to generate!")
        }
        if (!is.null(unmatched)){
            message("There are ", length(unmatched), " row headings do not match ",
                    "the database. Please verify it in the 'unmatched.csv'.")
            write.csv(sort(unmatched), paste0(filepath, "unmatched.csv"))
        }
        message("Total minutes used: ", as.numeric(Sys.time() - time.begin, 
                                                   units="mins"))
    }
}

#' @export
#' @rdname bindTXTs
bind_txts <- bindTXTs

#' @importFrom stringr str_replace_all
mapTxt2Df <- function(txt, ruleDf){
    # Map txt to data.frame
    # Based on rule.sheet template
    stopifnot(all(c("Pattern", "Extract", "Fld") %in% names(ruleDf)))
    origTxt <- scan(file=txt, what="", sep="\n", encoding="UTF-8", quiet=TRUE)
    origTxt <- paste(origTxt, collapse = "\n")[[1]]
    origTxt <- str_replace_all(origTxt, "\\\n{2,}|\\\f{1,}", "\\\n")
    
    isPtn <- unname(sapply(ruleDf$Pattern, function(ptn) grepl(ptn, origTxt)))
    rec <- rep(NA, nrow(ruleDf))
    rec[isPtn] <- mapply(function(ptn, rpl, str=origTxt){
        sub(ptn, rpl, origTxt)},
        ptn=ruleDf$Pattern[isPtn], rpl=ruleDf$Extract[isPtn]
    )
    names(rec) <- ruleDf$Fld
    return(rec)
}

#' @importFrom stringr str_replace_all str_replace str_split
getTxtHead <- function(txt, ruleDf){
    # List all the beginning words in the txt
    # that are not stored in rule.sheet template
    stopifnot(all(c("Pattern", "Extract", "Fld") %in% names(ruleDf)))
    origTxt <- scan(file=txt, what="", sep="\n", encoding="UTF-8", quiet=TRUE)
    origTxt <- paste(origTxt, collapse = "\n")[[1]]
    origTxt <- str_replace_all(origTxt, "\\\n{2,}|\\\f{1,}", "\\\n")
    origTxt <- str_replace_all(origTxt, "^\\s+", "")
    ptns <- str_replace(ruleDf$Pattern, "^[^\\u533b]*\\u533b\\u751f(.+)$", "\\1")
    ptns <- ptns[!grepl("^\\^", ptns)]
    ptns <- unique(str_replace_all(ptns, "^\\.\\+\\?*(\\\\n)*", "\\^"))
    heads <- unlist(str_split(origTxt, "\\n"))
    heads <- unique(str_replace_all(
        heads, "^([^\\s]+?)\\s.+$", "\\1"))
    i <- sapply(heads, function(exp) {
        ! any(sapply(ptns, function(ptn) grepl(ptn, exp)))
    })
    return(heads[i])
}
