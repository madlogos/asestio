#' Find the Variable Row
#'
#' Find the variable row and return a list of the index, row names and a data
#' frame containing pct_valid_val, pct_unique_val, pct_complete_val.
#' It makes use of the private function \pkg{aseshms}:::\code{checkTopRows}.
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param files You can manually assign the files. If null, you will select them
#' using a wizard.
#'
#' @return A list \describe{
#' \item{rowindex}{row index}
#' \item{vars}{col names (table head)}
#' \item{checkrows}{a data.frame containing pct_valid_val, pct_unique_val, pct_complete_val}
#' }
#' @export
#' @importFrom gWidgets2 gfile
#' @importFrom compiler cmpfun
#' @importFrom stringi stri_conv
#' @importFrom openxlsx read.xlsx getSheetNames
#'
#' @examples
#' \dontrun{
#' findVarRow()
#' }
findVarRow <- function(files=NULL){    
    #Sys.setlocale("LC_CTYPE","Chs")
    #addRtoolsPath()
    # get a file
    if (is.null(files) || ! file.exists(files)){
        files <- gfile(text = "Select xls files...", type = "open",
                       initial.dir = getOption("init.dir"),
                       filter = list('xls files'=list(patterns=c('*.xls?', '*.xls'))),
                       multi=FALSE, toolkit=guiToolkit(getOption("guiToolkit")))
        files <- stri_conv(files, "CP936", "UTF-8")
        stopifnot(file.exists(files))
    }
    sheets <- enc2native(getSheetNames(files))
    
    # calc
    l <- checkTopRows(files=files, sheets=sheets)
    names(l) <- sheets
    l <- l[!sapply(l, is.null)]
    rowindex <- sapply(l, function(l) {
        if (is.null(l)) return(NULL) else return(l[1, "RowNum"])
    })
    rowindex <- unlist(rowindex)
    vars <- sapply(names(rowindex), cmpfun(function(s, file=files, index=rowindex){
        df <- read.xlsx(file,s,colNames = FALSE, skipEmptyRows = FALSE,
                        detectDates=TRUE)
        return(unlist(df[rowindex[s],], use.names=FALSE))
    }))
    return(list(rowindex=rowindex, vars=vars, checkrows=l))
}

#' @export
#' @rdname findVarRow
find_var_row <- findVarRow

pVarInRow <- function(x, na.rm=FALSE){
	# percent of var-like values in a row
	if (length(x) > 0){
		x <- as.character(x)
		n <- is.na(as.numeric(x)) & !is.na(x) & nchar(x)<64
		return(sum(n)/(if (na.rm) sum(!is.na(x)) else length(x)))
	}else{
		return(NA)
	}
}
pUniqueInRow <- function(x, na.rm=FALSE){
	# percent of unique values in a row
	if (length(x) > 0) {
		x <- as.character(x)
		s <- is.na(as.numeric(x)) & !is.na(x) & nchar(x) < 64
		return(length(unique(x[s])) / (if (na.rm) sum(!is.na(x)) else length(x)))
	}else{
		return(NA)
	}
}

#' @importFrom openxlsx read.xlsx
#' @importFrom compiler cmpfun
checkTopRows <-function(files, sheets, rows=10){
	# return a df of pct_valid_val, pct_unique_val, pct_complete_val in top rows
	dfs <- lapply(sheets, function(s) {
		read.xlsx(file, s, colNames=FALSE, skipEmptyRows=FALSE,
				  detectDates=TRUE)})
	checkDf <- cmpfun(function(df, rows=rows){
		if (!is.null(df) && nrow(df)>0){
			out <- cbind(seq_len(rows), apply(df[seq_len(rows), ],
											  1, pVarInRow),
						 apply(df[seq_len(rows),], 1, pUniqueInRow),
						 apply(df[seq_len(rows),], 1,
							   function(x) sum(!is.na(x))/length(x)))
			out <- as.data.frame(out)
			names(out) <- c("RowNum", "Valid", "Unique", "Complete")
			out <- out[order(-out$Valid, -out$Unique,
							 -out$Complete, out$RowNum),]
		}else{
			out <- NULL
		}
	})
	lapply(dfs, checkDf, rows=rows)
}