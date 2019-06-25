#' asestio: An R Data I/O & Transformation Toolkit for ASES
#'
#' An analytics toolkit comprising of a series of work functions on data I/O and transformation.
#' The toolkit has a lot of featured funcionalities specifically designed for ASES HMS team.
#'
#' @details The toolkit comprises of several types of functions: \describe{
#' \item{File conversion}{\code{\link{conv_pdfs}()}, \code{\link{xls2xlsx}()},
#'   \code{\link{doc2docx}()}}
#' \item{Merge datasets}{\code{\link{bind_pdfs}()}, \code{\link{bind_txts}()},
#'   \code{\link{bind_excels}()}}
#' \item{Merge data rows/columns}{\code{\link{bind_cols}()}, \code{\link{bind_rows}()}}
#' \item{Other}{\code{\link{rehead_htmltbl}()}, \code{\link{simp_summ}()},
#'   \code{\link{find_var_row}()}}
#' }
#' @author \strong{Maintainer}: Yiying Wang, \email{wangy@@aetna.com}
#'
#' @importFrom magrittr %>%
#' @export %>%
#' @seealso \pkg{\link{recharts}}
#' @docType package
#' @keywords internal
#' @name asestio
NULL

#' @importFrom aseskit addRtoolsPath
.onLoad <- function(libname, pkgname="asestio"){

    if (Sys.info()[['sysname']] == 'Windows'){
        Sys.setlocale('LC_CTYPE', 'Chs')
    }else{
        Sys.setlocale('LC_CTYPE', 'zh_CN.utf-8')
    }
    if (Sys.info()[['machine']] == "x64") if (Sys.getenv("JAVA_HOME") != "")
        Sys.setenv(JAVA_HOME="")

    addRtoolsPath()

    # pkgenv is a hidden env under pacakge:asestio
    # -----------------------------------------------------------
    assign("pkgenv", new.env(), envir=parent.env(environment()))

    # ----------------------------------------------------------

    # options
    assign('op', options(), envir=pkgenv)
    options(stringsAsFactors=FALSE)

    pkgParam <- aseskit:::.getPkgPara(pkgname)
    toset <- !(names(pkgParam) %in% names(pkgenv$op))
    if (any(toset)) options(pkgParam[toset])
}

.onUnload <- function(libname, pkgname="asestio"){
    op <- aseskit:::.resetPkgPara(pkgname)
    options(op)
}


.onAttach <- function(libname, pkgname="asestio"){
    ver.warn <- ""
    latest.ver <- getOption(pkgname)$latest.version
    current.ver <- getOption(pkgname)$version
    if (!is.null(latest.ver) && !is.null(current.ver))
        if (latest.ver > current.ver)
            ver.warn <- paste0("\nThe most up-to-date version of ", pkgname, " is ",
			                   latest.ver, ". You are currently using ", current.ver)
    packageStartupMessage(paste("Welcome to", pkgname, current.ver,
                                 ver.warn))
}

