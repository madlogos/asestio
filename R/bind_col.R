#' Combine Similar Columns Into One
#'
#' Combine similar columns into one column within one spreadsheet.
#'
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param mapping.file Name of the field name mapping dictionary. Default
#' 'BindFieldNames.xlsx' in \code{toolkit} directory.
#' @param toolkit Directory of the dictionary.
#' @param sheet the sheet (number or name character) of the target workbook to
#' read in. Default 1.
#' @param output.format csv or xlsx. Default xlsx.
#'
#' @return A list. \itemize{
#'  \item \code{Output Results} shows the path where bind_output.xlsx is;
#'  \item \code{Dropped columns} shows the variables dropped during the combination.
#' }
#' @export
#' @importFrom gWidgets2 gfile
#' @importFrom compiler cmpfun
#' @importFrom readr read_csv write_csv
#' @importFrom readxl read_excel excel_sheets
#' @importFrom stringi stri_conv
#' @importFrom stringr str_detect
#' @importFrom openxlsx getSheetNames readWorkbook read.xlsx write.xlsx createStyle
#'
#' @examples
#' \dontrun{
#' bindColumns()
#' }
bindColumns <- function(mapping.file=paste0(toolkit, 'BindFieldNames.xlsx'),
                        toolkit=getOption("toolkit.dir"), sheet=1,
                        output.format=c('xlsx', 'csv')){
    # ------------Env codes no longer needed---------------
    # Sys.setlocale("LC_CTYPE", "Chs")  # Very important!
    # addRtoolsPath()

    output.format <- match.arg(output.format)

    #------Field name dictionary-------
    sources <- enc2native(getSheetNames(mapping.file))
    ynAssignMap <- .tcltkSelectItem(sources, "Select the mapping source.")
    if (!ynAssignMap[[1]]) {
        stop("No source designated!")
    }else {
        #map.source <- iconv(ynAssignMap[[2]], "UTF-8", "CP936")
        map.source <- enc2native(ynAssignMap[[2]])
        map <- readWorkbook(mapping.file, map.source)
        map <- map[order(map$InDb, map$InFile), ]
        map[, 1] <- toupper(map[, 1])
        map[, 2] <- toupper(map[, 2])

        # read files--------------------
        if (Sys.info()['sysname'] == "Windows"){
            raw.file <- invisible(choose.files(
                paste0(getOption("init.dir"), "*.*"), multi=FALSE,
                caption="Select the raw data file ...",
                filters=rbind(matrix(c("Excel files (*.xls?)", "*.xls?;*.xls",
                                       "csv files (*.csv)", "*.csv"),
                                     byrow=TRUE, nrow=2),
                              Filters["All",])))
        }else{
            raw.file <- gfile("Select the raw data file ...", type='open',
                          initial.dir=getOption("init.dir"),
                          filter=list('xls* files'=list(
                              patterns=c('*.xls?', '*.xls')),
                              'csv files'=list(patterns=c('.csv'))),
                          multi=FALSE)
            raw.file <- stri_conv(raw.file, "CP936", "UTF-8")
        }

        if (!file.exists(raw.file)) stop("No file designated!")
        vld_sheets <- excel_sheets(raw.file)
        if (is.character(sheet)) if (! sheet %in% vld_sheets)
            stop("sheet is not a valid sheet name in the workbook you selected.")
        dta <- read_excel(raw.file, sheet=sheet)
        input.cols <- names(dta)
        if (str_detect(raw.file, "csv$"))
            dta <- read_csv(raw.file)

        #----------bind columns, need to recode-----------------------
        maps <- split(map, map$InDb)
        bind.all <- unlist(lapply(maps, function(x){
                all(x$BindAll) | !(any(x$BindAll))
            }))
        if (! all(bind.all))
            stop(cat("BindAll not consistent:\n", names(bind.all)[!bind.all]))

        origColNames <- data.frame(InFile=names(dta))
        origColNames <- merge(origColNames, map, by="InFile",
                              all.x=TRUE, sort=FALSE)
        origColNames$InDb[is.na(origColNames$InDb)] <-
            origColNames$InFile[is.na(origColNames$InDb)]
        origColNames <- c(
            "FILE", "SHEET", as.character(
                origColNames$InDb[! duplicated(origColNames$InDb)]))

        dta <- dta[, ! names(dta) %in% map$InFile[is.na(map$InDb)]]

        listbindVal <- lapply(maps, cmpfun(funBindList), dat=dta)
        class.vars <- vapply(listbindVal, class, FUN.VALUE=character(1L))
        listbindVal[class.vars=='NULL'] <- NULL
        df <- as.data.frame(listbindVal, stringsAsFactors=FALSE)
        dta <- cbind(dta[, !names(dta) %in% c(map$InDb, map$InFile)], df)

        dta <- dta[, intersect(origColNames, names(dta))]
        # for (i in seq_len(nrow(map))){
        #     if (is.na(map$InDb[i])){
        #         dta[,c(map$InFile[i])] <- NULL   #drop
        #     }else if(map$InFile[i] != map$InDb[i]){
        #         if (map$InFile[i] %in% names(dta)){
        #             if (!map$InDb[i] %in% names(dta)){
        #                 names(dta) <- sub(map$InFile[i],map$InDb[i],names(dta))
        #             }else{
        #                 if (map$BindAll[i]){            #BindAll=T bind all contents
        #                     dta[,map$InDb[i]] <- paste(dta[,map$InDb[i]],
        #                                                dta[,map$InFile[i]],sep=";")
        #                     dta[,map$InDb[i]] <-
        #                         sub("NA;|;NA$|^[[:space:]]*NA[[:space:]]*$","",
        #                                              dta[,map$InDb[i]])
        #                 }else if (!map$BindAll[i]){     #BindAll=F bind vacancies
        #                     if (!map$InDb[i] %in% names(dta)) dta[,map$InDb[i]] <- NA
        #                     dta[is.na(dta[,map$InDb[i]]),map$InDb[i]] <-
        #                         dta[is.na(dta[,map$InDb[i]]),map$InFile[i]]
        #                 }
        #                 dta[,map$InFile[i]] <- NULL
        #             }
        #         }
        #     }
        # }
        if (length(raw.file) > 1)
            message("You selected multiple files, we will pick up the first one's directory.")
        raw.path <- str_replace_all(raw.file[1], "\\\\", "/")
        raw.path <- str_replace(raw.path, "^(.+/)[^/]+$","\\1")

        if (! str_detect(raw.path, ".+\\\\$|.+/$")) raw.path <- paste0(raw.path, "/")
        if (output.format == 'csv'){
            write.csv(dta, paste0(raw.path, "bind_output.csv"), na="")
        }else if (output.format=='xlsx'){
            write.xlsx(dta, file=paste0(raw.path, "bind_output.xlsx"),
                       sheetName="Sheet1",
                       headerStyle=createStyle(fgFill="#E8E8E8",
                                               fontName='Arial Narrow'),
                       withFilter=TRUE)
        }

        drop.vars <- data.frame(
            idx=which(!input.cols %in% c("FILE", "SHEET", map$InDb, map$InFile)),
            var=input.cols[! input.cols %in% c("FILE", "SHEET", map$InDb,
                                               map$InFile)]
            )
        return(list("Output Result"=paste0("The cleaned dataset 'bind_output.",
                                    output.format, "' is in the folder ",
                                    raw.path),
                    "Dropped columns"=drop.vars))
    }
}

#' @export
#' @rdname bindColumns
bind_columns <- bindColumns


funBindList <- function(x, dat){
    # loop call .funBindCols to bind columns by grouped columns
    # Args:
    #    x: mapping df
    #    dat: a data.frame
    # Return:
    #    Binded df
    stopifnot(all(c('InFile', 'InDb', 'BindAll') %in% names(x)))

    if (any(names(dat) %in% c(x$InFile, x$InDb))){
        .funBindCols(dat[, intersect(names(dat), c(x$InFile, x$InDb))],
                     bind.all=x$BindAll[1])
    }
}

#' @importFrom stringr str_detect str_replace_all
.funBindCols <- function(x, bind.all=FALSE){
    # Base function to bind grouped columns (like cbind) into one
    # Args:
    #   x: object to combine
    #   bind.all: logical. If TRUE, bind all contents; if FALSE, extract the first
    #            non-empty content
    # Return:
    #   binded vector based on x
    if (is.null(dim(x))){
        v <- x
    }else{
        x <- as.data.frame(sapply(as.list(x), trimws), stringsAsFactors=FALSE)
        x[is.na(x)] <- ""
        v <- do.call('paste', c(x, sep="%&%"))
        v <- str_replace_all(v, "^(%&%)+|NA(%&%)+|(%&%)+NA|(%&%)+$", "")
        if (bind.all){
            v <- str_replace_all(v, "(%&%)+", ";")
        }else{
            v <- str_replace_all(v, "^(.*?)%&%.*$", "\\1")
        }
        v[str_detect(v, "^$|^NA$")] <- NA
    }
    return(v)
}
