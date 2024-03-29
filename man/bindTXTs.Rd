% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/merge_txt.R
\name{bindTXTs}
\alias{bindTXTs}
\alias{bind_txts}
\title{Combine Txts And Convert Them to A Data Frame}
\usage{
bindTXTs(rule.file = paste0(path, "IndRptToTable(RegExp).xlsx"),
  path = getOption("toolkit.dir"), txts = NULL, use.parallel = TRUE)

bind_txts(rule.file = paste0(path, "IndRptToTable(RegExp).xlsx"),
  path = getOption("toolkit.dir"), txts = NULL, use.parallel = TRUE)
}
\arguments{
\item{rule.file}{Full path of the conversion dictionary file. Default
\code{paste(path, dict name)}.}

\item{path}{Path of the directory to the dictionary file.}

\item{txts}{Txts assigned for combination. If NULL, a wizard will guide you
to assign them.}

\item{use.parallel}{Logical, whether to apply parallel computation when there 
are more than 50 files to process. Default \code{TRUE}.}
}
\value{
Nothing
}
\description{
Combine Txts derived from PDFs and convert them to a data frame using a
conversion dictionary stored in IndRptToTable(RegExp).xlsx.
}
\details{
The function yields two files: an .xlsx file containing merge data and a .csv
containing the unmatched line headers that are potentially new variables.
}
\examples{
\dontrun{
bindTXTs()
}
}
\author{
Yiying Wang, \email{wangy@aetna.com}
}
