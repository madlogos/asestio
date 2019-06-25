#' Combine Excel Files
#'
#' Combine several excel files into one.
#'
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param sheet.num Sheet number where data is stored. If null,
#'   then \pkg{\link{dplyr}}::\code{\link[dplyr]{bind_rows}} all the sheets.
#' @param files Designate the files yourself, or manually select them.
#' @param output.format csv or xlsx. Default xlsx.
#'
#' @return A list containing a data frame listing the files and sheets combined.
#' @export
#' @importFrom gWidgets2 gfile
#' @importFrom compiler cmpfun
#' @importFrom foreach foreach %dopar%
#' @importFrom parallel makeCluster stopCluster
#' @importFrom doParallel registerDoParallel stopImplicitCluster
#' @importFrom readxl read_excel
#' @importFrom readr write_csv write_excel_csv
#' @importFrom openxlsx read.xlsx write.xlsx createStyle getSheetNames
#' @importFrom dplyr bind_rows
#' @importFrom stringr str_detect str_replace_all
#' @importFrom stringi stri_conv
#' @seealso \code{\link[dplyr]{bind_rows}} in package \pkg{\link{dplyr}}.
#' @examples
#' \dontrun{
#' bindExcels()
#' }
bindExcels <- function(sheet.num=1,
                       files=NULL, output.format=c('xlsx', 'csv')){
    #Sys.setlocale("LC_CTYPE","Chs")   # Very important!
    #options(java.parameters = "-Xmx1024m"

    # -------ensure Rtools is included in Sys.Path-------
    #addRtoolsPath()
    # -------bind datasets--------------
    #pdfname <- as.data.frame(list.files(filepath,pattern=".+.pdf"))
    #wbname <- as.list(list.files(filepath,pattern=".+.xls.*"))
    output.format <- match.arg(output.format)

    # ----------------------

    if (!is.null(files)){
        files <- unlist(files)
        wb <- files[str_detect(files, "[Xx][Ll][Ss]$|[Xx][Ll][Ss][MmXxBb]$") &&
                        file.exists(files)]
    }else{
        if (Sys.info()['sysname'] == "Windows"){
            wb <- invisible(choose.files(
                paste0(getOption("init.dir"), "*.*"),
                caption="Select Excel files...",
                filters=rbind(matrix(c("Excel files (*.xls?)", "*.xls?;*.xls"), nrow=1),
                              Filters["All",])))
        }else{
            wb <- gfile("Select .xls* files...", type='open',
                    initial.dir=getOption("init.dir"),
                    filter=list('xls* files'=list(patterns=c('*.xls?', '*.xls'))),
                    multi=TRUE, toolkit=guiToolkit(getOption("guiToolkit")))
            wb <- stri_conv(wb, "CP936", "UTF-8")
        }
        if (! all(file.exists(wb)))
            warning(paste("A total of", sum(! file.exists(wb)), "workbooks not existing."))
        wb <- wb[file.exists(wb)]
    }
    if (is.null(wb) || length(wb)==0) stop("No file designated!")
    if (is.null(sheet.num))
        invisible(paste("You assigned sheet.num = NULL,",
                        "so all the sheets contained in the files will be combined.\n",
                        "dplyr::bind_rows will be called."))

    ## ---------UDF to rm all NA cols----------
    if (length(wb) > 0){
        bindSheetCompress <- cmpfun(.bindSheetCompress)

        if (length(wb) < 20){
            lstRawDta <- lapply(wb, bindSheetCompress, sht=sheet.num,
                                compress=TRUE)
        }else{
            ## cluster computation
            cl <- makeCluster(detectCores(logical=FALSE), type="PSOCK")
            registerDoParallel(cl, detectCores(logical=FALSE))
            lstRawDta <- invisible(
                foreach(i = 1:length(wb)) %dopar% bindSheetCompress(
                    wb[i], sht=sheet.num, compress=TRUE))
            stopImplicitCluster()
            stopCluster(cl)
        }
    }else{
        message("No sheets for bind!")
        return(NULL)
    }

    # in lstRawDta, each list contains two lists, nRows and data.frame
    lstNRows <- lapply(lstRawDta, cmpfun(function(L) L[[1]]))  # get list of nRows
    lstDta <- lapply(lstRawDta, cmpfun(function(L) L[[2]]))  # get list of data.frame
    nRows <- bind_rows(lstNRows)
    dta <- bind_rows(lstDta)
    # names(nRows) <- c("FILE", "SHEET", "Input")

    dtVar <- grep(enc2native("\u65e5\u671f$|\u65f6\u95f4$"), names(dta))
    if (length(dtVar) > 0) {
        lstdtVar <- lapply(dtVar, cmpfun(.funConvDate), data=dta)
        dta[, dtVar] <- do.call('cbind', lstdtVar)
        dta[, dtVar] <- lapply(dtVar, function(var){
            as.Date(dta[, var], orig="1970-1-1", tz="")
        })
        for (i in dtVar) class(dta[, i]) <- "Date"
    }
    df <- data.frame(lapply(dta, as.character), stringsAsFactors=FALSE)
    # df <- df[,cols]  # restore the order of the columns
    names(df) <- str_replace_all(names(df), "[\\. ]", "")

    # ------------Export data--------------------------
    #path <- str_replace_all(wb[1], "^(.+\\\\)[^\\]+\\.[Xx][Ll][Ss].{0,1}$", "\\1")
    path <- gsub("^(.+\\\\)[^\\]+\\.[Xx][Ll][Ss].{0,1}$", "\\1", wb[1])
    if (! str_detect(path, ".+\\\\$")) path <- paste0(path,"\\\\")
    if (output.format=="xlsx"){
        write.xlsx(df, file=paste(path, "bind.xlsx", sep=""),
                     sheetName="Sheet1",
                     headerStyle=createStyle(fgFill="#E8E8E8",
                                             fontName='Arial Narrow'),
                     withFilter=TRUE)
    }else if (output.format=='csv'){
        write.csv(df, paste0(filepath,"bind.csv"), na="",
                  fileEncoding = "UTF-8")
    }
    message(cat("The combined dataset 'bind.", output.format,
                "' is in the folder", path, "!\n"))
    nRowsOut <- dcast(df, FILE+SHEET~., length, value.var="FILE")
    names(nRowsOut) <- c("FILE", "SHEET", "Output")
    return(list("Result of Combination"=merge(
        nRows, nRowsOut, by=c("FILE","SHEET"), all.x=TRUE, sort=FALSE)))
}

#' @export
#' @rdname bindExcels
bind_excels <- bindExcels

#' @importFrom stringr str_detect
.funConvDate <- function(var, data){
    # Convert data[, var] to date
    # Args:
    #   var: var name
    #   data: data.frame
    #
    # Return:
    #   Data frame with column(var) converted to dates
    v <- data[, var]
    t <- rep(NA, length(v))
    if (is.numeric(v)) v <- as.character(v)
    if (is.character(v)){
        t[str_detect(v, "^\\d{5}$|^\\d{5}\\.\\d{1,}")] <-
            as.POSIXct(60*60*24* as.numeric(
                v[str_detect(v, "^\\d{5}$|^\\d{5}\\.\\d{1,}")]),
                origin="1899-12-30")
        t[str_detect(v, "^\\d{8}$")] <-
            as.POSIXct(as.Date(v[str_detect(v, "^\\d{8}$")], format="%Y%m%d", tz=""))
        t[is.na(t)] <- try(as.POSIXct(v[is.na(t)]), silent=TRUE)
        return(as.Date(as.POSIXct(as.numeric(t), origin="1970-1-1"), tz=""))
    } else {
        return(as.Date(v, tz=""))
    }
}

#' @importFrom dplyr bind_rows
#' @importFrom stringr str_split
#' @importFrom readxl read_excel excel_sheets
.bindSheetCompress <- function(file, sht, compress=TRUE, ...){
    # Combine sheets from file and remove all-NA columns
    # Args:
    #   file: xlsx file
    #   sht: Excel sheet
    #   compress: Logical. If TRUE, remove all-NA columns.
    # Return:
    #   A list of two data.frames: nRows and combined data.frame
    #   Col names are capitalized. dots in var names are removed.

    #wbname <- str_replace_all(file, ".+\\\\([^\\]+$)", "\\1")
    if (length(file) > 1) {
        warning("You provided more than one file. Only the first one will be used.")
        file <- file[[1]]
    }
    file.pieces <- unlist(str_split(file, "\\\\|/"))
    wbname <- file.pieces[length(file.pieces)]

    if (is.null(sht)) {
        sht <- enc2native(excel_sheets(file))
        st <- lapply(sht, function(sht.name) {
            df <- read_excel(enc2native(file), sheet=sht.name)
            if (nrow(df)>0) {
                df$file <- enc2native(wbname)
                df$sheet <- sht.name
            }else{
                df <- data.frame(df, file=character(0), sheet=character(0))
            }
            return(df)
        })
        cols <- vapply(st, names, FUN.VALUE=character(1L))
        cols <- unlist(cols)
        st <- bind_rows(st)
    } else {
        st <- read_excel(enc2native(file), sht)
        if (!is.null(st)) {
            sheet.name <- enc2native(excel_sheets(file))
            if (nrow(st)>0){
                st$file <- enc2native(wbname)
                cols <- names(st)
                st$sheet <- sheet.name[sht]
            }else{
                return(list(
                    "Data Source"=data.frame(
                        FILE=wbname, SHEET=sheet.name[sht], Input=0,
                        stringsAsFactors=FALSE),
                    "Combined Data"=NULL))
            }
        }else{
            return(list("Data Source"=NULL, "Combined Data"=NULL))
        }
    }
    # remove duplicated col names and superassign it to outer environ
    # cols <<- cols[!duplicated(cols)]
    # empty value <- NA
    st.names <- names(st)
    st <- sapply(st.names, function(var) {
        st[, var][grepl("^$", st[, var])] <- NA
        return(st[, var])
    })
    names(st) <- st.names
    if (is.null(dim(st))) st <- as.list(st)
    st <- data.frame(st, stringsAsFactors=FALSE)
    n_na <- colSums(as.data.frame(apply(st, 2, is.na)))
    rm_cols <- which(n_na==nrow(st))
    if (compress & length(rm_cols)>0)
        st <- st[, -which(n_na==nrow(st))]  # rm all NA cols
    names(st) <- toupper(str_replace_all(names(st), "[\\. ]", ""))
    tabSource <- dcast(st, FILE+SHEET~., length, value.var="FILE")
    names(tabSource) <- c("FILE", "SHEET", "Input")

    if (any(duplicated(names(st)))){
        stop(paste("Duplicated var names detected.\nFile: ", file,
                   "\nSheet: ", sht, "\nDuplicated Var: ",
                   paste(names(st)[duplicated(names(st))], collapse=", ")))
    }
    return(list("Data Source"=tabSource, "Combined Data"=st))
}
