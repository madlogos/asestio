#-------table format-----------
#' Reformat HTML Table
#'
#' Convert a data frame to an HTML table object and reformat it.
#'
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param dataset The dataset to draw table.
#' @param heading The heading you want to input.
#'        '|' indicates colspan, '=' indicates rowspan.
#' @param footRows The last several rows as <tfoot>.
#' @param align Alignment of columns.
#' @param concatCol Index of columns to concatenate,
#'        to make the table look hierachical.
#' @param caption Table caption.
#' @param tableWidth Width of the table.
#'
#' @return A table in HTML.
#' @export
#' @importFrom knitr kable
#' @importFrom stringr str_replace_all str_replace
#' @seealso \code{\link{kable}}
#'
#' @examples
#' \dontrun{
#' ## A reformatted table with colspan/colrow=2
#' heading <- matrix(c("Sepal", NA, "Petal", NA, "Species", "Length", "Width",
#'                     "Length", "Width", NA), byrow=TRUE, nrow=2)
#' reheadHTMLTable(head(iris), heading)
#' }
reheadHTMLTable <- function(dataset, heading, footRows=0,
                           align=c('left', rep('center', ncol-1)),
                           concatCol=NULL, caption=NULL,
                           tableWidth='100%'){
    if ((!is.null(dataset) & !is.data.frame(dataset)) |
        !(is.data.frame(heading) | is.matrix(heading) | is.vector(heading))){
        stop(paste0('`dataset` must be a data.frame, while you gave a ',
                    class(dataset),
                    '\n`heading` must be a vector/matrix/data.frame, ',
                    'while you gave a ', class(heading),"."))
    }else{
        if (is.vector(heading)) heading <- t(matrix(heading))
        if (!is.null(dataset)){
            ncol <- ncol(dataset)
            if (ncol!=ncol(heading))
                stop(paste("Not equal counts of columns! Dataset has",
                           ncol, "cols, while heading has",ncol(heading), '.'))
        }else{
            ncol <- str_replace(str_replace_all(htmltable, '.+?<t([dhr]).+?', '\\1'),
                                '(^[dhr]+?)[^dhr].+$','\\1')
            ncol <- table(strsplit(ncol, "")[[1]])
            ncol <- floor((ncol[['h']]+ncol[['d']])/ncol[['r']])
        }
        align_simp <- substr(tolower(align),1,1)
        if (!all(align_simp %in% c('l','c','r'))){
            stop('`align` only accepts values of "l(eft)", "c(enter)" and "r(ight)".')
        }else{
            align[align_simp=="l"] <- "left"
            align[align_simp=="c"] <- "center"
            align[align_simp=="r"] <- "right"
        }
        if (length(align) > ncol(heading)){
            align <- align[1:ncol(heading)]
        }else if (length(align)<ncol(heading)){
            align <- c(align[1:length(align)],
                       rep(align[length(align)], ncol(heading)-length(align)))
        }
        align_simp <- substr(tolower(align),1,1)
        # loadPkg('knitr')

        dataset <- as.data.frame(dataset)
        if (!is.null(concatCol)){
            for (icol in concatCol){
                col <- as.character(dataset[,icol])
                lag <- c(NA, as.character(dataset[1:(nrow(dataset)-1), icol]))
                col[col==lag] <- ""
                dataset[, icol] <- col
            }
        }

        if (!(is.null(footRows) | footRows==0)){
            if (footRows>=nrow(dataset))
                stop("footRows cannot be >= number of datatable rows!")
            htmlBody <- kable(
                dataset[1:(nrow(dataset)-footRows), ], format='html',
                align=align_simp, row.names=FALSE)
            htmlFoot <- kable(
                dataset[(nrow(dataset)-footRows+1): nrow(dataset),],
                format='html', align=align_simp, row.names=FALSE)
            htmlBody <- str_replace_all(htmlBody, "(^.+</tbody>).+$", "\\1")
            htmlFoot <- str_replace_all(htmlFoot, "^.+<tbody>(.+)</tbody>.+$", 
                                        "<tfoot>\\1</tfoot>")
            htmltable <- paste0(htmlBody,"\n", htmlFoot, "\n</table>")
        }else{
            htmltable <- kable(
                dataset, format='html', align=align_simp, row.names=FALSE)
        }

        if (!is.null(caption)){
            htmltable <- str_replace_all(
                htmltable, "<table>", paste0("<table>\n<caption>", caption, "</caption>"))
        }
        class(htmltable) <- 'knitr_kable'
        attributes(htmltable) <- list(format='html', class='knitr_kable')
        rehead <- '<thead>'
        for (j in 1:ncol(heading)){
            if (all(is.na(heading[, j]))){
                heading[1,j] <- '$'
                if (nrow(heading)>1) heading[2:nrow(heading),j] <- "|"
            }
        }
        heading[1,][is.na(heading[1,])] <- "="
        if (nrow(heading)>1){
            heading[2:nrow(heading), ][is.na(heading[2:nrow(heading), ])] <- "|"
        }
        dthead <- heading
        for (i in 1:nrow(heading)){
            for (j in 1:ncol(heading)){
                dthead[i,j] <- ifelse(
                    heading[i, j] %in% c('|','='), "", paste0(
                        '   <th style="text-align:', align[j], ';"> ',
                        heading[i,j], ' </th>\n'))
                if (! heading[i,j] %in% c("|","=")){
                    if (i==1 & heading[i,j]=="$"){
                        dthead[i,j] <- paste0(
                            '   <th rowspan="', nrow(heading),
                            '" style="text-align:', align[j],
                            ';">&nbsp;&nbsp;&nbsp;</th>\n')
                    }
                    if (j < ncol(heading)) {
                        if (heading[i, j+1] == "="){
                            colspan <- paste0(
                                heading[i, (j+1):ncol(heading)], collapse="")
                            ncolspan <- nchar(sub("^(=+).*$","\\1",colspan))+1
                            dthead[i,j] <- sub(
                                '<th ', paste0('<th colspan="', ncolspan,'" '),
                                               dthead[i, j])
                            dthead[i,j] <- sub(
                                'align: *?(left|right)', paste0('align:center'),
                                dthead[i,j])
                        }
                    }
                    if (i < nrow(heading)){
                        if (heading[i+1, j] == "|"){
                            rowspan <- paste0(
                                heading[(i+1):nrow(heading), j], collapse="")
                            nrowspan <- nchar(sub("^(\\|+).*$", "\\1",rowspan))+1
                            if (str_detect(dthead[i, j], "colspan")){
                                if (sum(!heading[i:(i+nrowspan-1), j:(j+ncolspan-1)]
                                        %in% c('=','|'))==1){
                                    dthead[i,j] <- sub(
                                        '<th ', paste0('<th rowspan="', nrowspan,'" '),
                                        dthead[i, j])
                                }else{
                                    dthead[i,j] <- sub('colspan.+?style', paste0(
                                        'rowspan="', nrowspan,'" style'),
                                        dthead[i, j])
                                }
                            }else{
                                dthead[i,j] <- sub('<th ', paste0(
                                    '<th rowspan="', nrowspan,'" '), dthead[i, j])
                            }
                        }
                    }
                }
            }
        }
        for (i in 1:nrow(heading)){
            rehead <- paste0(rehead, '\n  <tr>\n', paste0(dthead[i, ], collapse=""),
                             '  </tr>\n', collapse="")
        }
        rehead <- paste0(rehead,'</thead>')
        rehead <- str_replace_all(htmltable, '<thead>.+</thead>', rehead)
        class(rehead) <- class(htmltable)
        attributes(rehead) <- attributes(htmltable)
        return(sub('<table', paste0('<table width=', as.character(tableWidth)),
                   rehead))
    }
}

#' @export
#' @rdname reheadHTMLTable
rehead_htmltbl <- reheadHTMLTable