#' Combine Rows Into One
#'
#' Combine rows into one which are actually sole record.
#'
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param df You can assign a data.frame yourself. If null, you will manually
#' select a spreadsheet and extract the data.
#' @param output.format xlsx or csv. Default xlsx.
#'
#' @return Nothing
#' @export
#' @importFrom compiler cmpfun
#' @importFrom readr read_csv
#' @importFrom readxl read_excel
#' @importFrom stringi stri_conv
#' @importFrom stringr str_detect
#' @importFrom openxlsx read.xlsx write.xlsx createStyle
#'
#' @examples
#' \dontrun{
#' bindRows()
#' }
bindRows <- function(df = NULL, output.format = c('xlsx', 'csv')){
    # Sys.setlocale("LC_CTYPE","Chs")  # Very important!
    # sapply(c("dplyr","openxlsx","gWidgets2","gWidgets2RGtk2","compiler"),
    #       require,character.only=TRUE)
    # options(guiToolkit="RGtk2")
    # addRtoolsPath()

    output.format <- match.arg(output.format)

    #----------read files--------------------
    if (is.null(df)){
        if (Sys.info()['sysname'] == "Windows"){
            raw.file <- invisible(choose.files(
                paste0(getOption("init.dir"), "*.*"), multi=FALSE,
                caption="Select the raw data file...",
                filters=rbind(matrix(c("Excel files (*.xls?)", "*.xls?;*.xls",
                                       "csv files (*.csv)", "*.csv"),
                                     byrow=TRUE, nrow=2),
                              Filters["All",])))
        }else{
            raw.file <- gfile("Select the raw data file...",type='open',
                          initial.dir=getOption("init.dir"),
                          filter=list('xls* files'=list(
                              patterns=c('*.xls?','*.xls')),
                              'csv files'=list(patterns=c('.csv'))),
                          multi=FALSE)
            raw.file <- stri_conv(raw.file, "CP936", "UTF-8")
        }

        if (!file.exists(raw.file)) stop("No file designated!")
        if (str_detect(raw.file, "[Xx][Ll][Ss]$|[Xx][Ll][Ss][XxMmBb]$")) {
            sheets <- enc2native(excel_sheets(raw.file))
            dims <- get_worksheet_dims(raw.file, sheets)
            sheet.sel <- guiselect_sheet(raw.file, sheets)
            if (!sheet.sel[[1]]) stop("You cancelled actions.")  # output = FALSE
            sheet.sel <- which(sheets==names(dims)[dims==sheet.sel[[2]]])
            dta <- read_excel(raw.file, sheet=sheet.sel)
        }
        if (str_detect(raw.file, "\\.[Cc][Ss][Vv]$"))
            dta <- read_csv(raw.file)
    }else{
        dta <- df
    }
    #ID var----------
    vars <- names(dta)
    var <- .funSelVar(vars)
    if (!var[[1]]) stop("You did not select any identifier!")
    else var <- var[[2]]
    #---------Combine rows-----------------
    dt <- split(dta, dta[, var])
    sn <- as.matrix(sapply(dt, function(x) nrow(x)))
    snNoBind <- row.names(sn)[sn[, 1]==1]
    snBind <- row.names(sn)[sn[, 1]>1]

    bind.all <- vapply(vars, cmpfun(.funBindMode), FUN.VALUE=logical(1L), dat=dta)

    after.bind <- as.data.frame(t(sapply(dt[snBind], cmpfun(funBindRows),
                                         bind.all=bind.all)))
    if (ncol(after.bind)>0) {
        names(after.bind) <- vars
        output <- rbind(do.call('rbind', dt[snNoBind]), after.bind)
    }else{
        output <- dt[[snNoBind]]
    }

    raw.path <- str_replace_all(raw.file[1], "^(.+\\\\)[^\\]+\\.[Xx][Ll][Ss].{0,1}$",
                     "\\1")
    if (! str_detect(raw.path, ".+\\\\$")) raw.path <- paste0(raw.path,"\\\\")
    if (output.format=='csv'){
        write.csv(output,paste0(raw.path,"bind_rows.csv"), na="")
    }else if (output.format=='xlsx'){
        write.xlsx(output,file=paste0(raw.path,"bind_rows.xlsx"),
                     sheetName="Sheet1",
                     headerStyle=createStyle(
                         fgFill="#E8E8E8",
                         fontName='Arial Narrow')
                     )
    }
    return(paste0("The cleaned dataset 'bind_rows.", output.format,
                  "' is in the folder ", raw.path))
}

#' @export
#' @rdname bindRows
bind_rows <- bindRows

#' @importFrom readxl read_excel
get_worksheet_dims <- function(file, sheets=NULL){
    # get dims in each sheet of a workbook
    # Args:
    #   file: workbook file
    #   sheets: sheets num or names
    # Return:
    #   named vector
    dims <- as.data.frame(vapply(sheets, function(x) {
        d <- try(dim(read_excel(raw.file, sheet=x)), silent=TRUE)
        if (is.null(d)) d <- c(0, 0) else d
    }, FUN.VALUE=numeric(2L)))
    # df --> named vec
    dims <- vapply(names(dims), function(x) {
        paste(c(x, paste(dims[, x], collapse=" ")), collapse=" ")
    }, FUN.VALUE=character(1L))
    return(dims)
}

#' @importFrom gWidgets2 gfile gwindow gvbox ggroup gframe gradio gbutton gaction
#' @importFrom gWidgets2 addHandlerChanged
#' @importFrom readxl read_excel excel_sheets
guiselect_sheet <- function(file, sheets=NULL){
    # GUI Select a sheet in a spreadsheet
    # Args:
    #   file: excel file
    #   sheets: the sheets to read, num or chr
    # return:
    #   list(output[bool], datasheet[df])

    # ----check args----
    stopifnot(file.exists(file) && grepl("xls[xbm]?$", tolower(file)))
    vld_sht <- enc2native(excel_sheets(file))
    if (is.null(sheets))
        sheets <- vld_sht

    if (! all(sheets %in% vld_sht)){
        warning("A total of ", sum(! sheets %in% vld_sht), " sheets not existing.")
        sheets <- sheets[sheets %in% vld_sht]
    }

    # ---get dims of each sheet---
    dims <- get_worksheet_dims(file, sheets=sheets)

    # GUI wizard
    window <- gwindow("Select the sheet", width=200, height=200)
    box <- gvbox(cont=window)
    addHandlerChanged(window, handler=function(...){
        gtkMainQuit()
    })
    gg1 <- ggroup(cont=box)
    gg2 <- ggroup(cont=box, horizontal = TRUE)
    box1 <- gvbox(cont=gg1)
    frm1 <- gframe("Sheet Name (nRow nCol):", cont=box1)
    chkmap <- gradio(items=dims, selected=1, index=TRUE, cont=frm1)
    box21 <- gvbox(cont=gg2)
    box22 <- gvbox(cont=gg2)
    actOK <- gaction("  OK  ", "OK",
                     handler=function(h,...){
                         dsheet <<- enc2native(svalue(chkmap))
                         output <<- TRUE
                         dispose(window)
                     })
    buttonOK <- gbutton(action=actOK, cont=box21)
    actCancel <- gaction("Cancel", "Cancel",
                         handler=function(h,...){
                             dsheet <<- NULL
                             output <<- FALSE
                             dispose(window)
                         })
    buttonCancel <- gbutton(action=actCancel, cont=box22)
    gtkMain()
    return(list(output, dsheet))
}

#' @importFrom stringr str_detect
.funBindMode <- function(col.name, dat){
    # judge if use bind.all mode
    # Args:
    #   col.name: name of the columns
    #   dat: df
    # Return:
    #   TRUE or FALSE

    stopifnot(is.data.frame(dat))
    stopifnot(col.name %in% names(dat))

    d <- dat[!is.na(dat[, col.name]), col.name]
    if (length(d)==0){
        return(FALSE)
    }else{
        has_qual_words <- str_detect(enc2native(d, "[\u62D2\u5F03\u9634\u9633\\<\\>\\+\\-]"))
        coerceNA <- try(as.numeric(d[! has_qual_words]), silent=TRUE)
        if (length(! has_qual_words)==0){
            p <- 0
        }else{
            p <- sum(is.na(coerceNA)) / length(! has_qual_words)
        }
        output <- (p>=0.5 || str_detect(tolower(col.name), enc2native("\u603B\u7ED3|\u5C0F\u7ED3?"))) &&
            ! str_detect(tolower(col.name), enc2native(paste0(
                "name|gender|sex|company|\u59D3\u540D|\u5355\u4F4D|\u90E8\u95E8|\u6027\u522B|",
                "department|\u8BC1$|\u53F7$|\u5EA6$|\u8054\u7CFB|\u4E2D\u5FC3$|\u7C7B\u578B$|provider|",
                "^\u5C3F|\u5EA6$|\u7EA7^|\u91CF$|\u6570$|\u503C$"))
            )
        return(output)
    }
}

#' @importFrom stringr str_replace_all str_detect
funBindRows <- function(x, bind.all){

    if (is.null(dim(x))){
        v <- x
    }else{
        v <- do.call('paste', c(as.data.frame(t(x)), sep="%&%"))
        v <- str_replace_all(v, "^%&%|NA%&%|%&%NA|%&%$","")
        v[bind.all] <- str_replace_all(v[bind.all], "%&%", ";")
        v[!bind.all] <- str_replace_all(v[!bind.all], "^(.*?)%&%.*$", "\\1")
        v[str_detect(v, "^$|^NA$")] <- NA
    }
    return(v)
}
