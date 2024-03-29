% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/conv_pdf.R
\name{convPDFs}
\alias{convPDFs}
\alias{conv_pdfs}
\title{Convert PDF to Text Files}
\usage{
convPDFs(mode = 1L, silent = FALSE, converter = c("pdftotext",
  "pdftohtml"), converter.path = NULL, use.parallel = TRUE)

conv_pdfs(mode = 1L, silent = FALSE, converter = c("pdftotext",
  "pdftohtml"), converter.path = NULL, use.parallel = TRUE)
}
\arguments{
\item{mode}{integer, Default 1L. \describe{
\item{1L}{as is (-raw)} \item{2L}{with layout (-layout)}
\item{3L}{no breaks (-nopgbrk)} \item{4L}{without format ()}
\item{5L}{as table (-table)} \item{6L}{simplified (-simple)}
}}

\item{silent}{logical, whether show the results of the conversion.
Default \code{FALSE}.}

\item{converter}{which executable program to use to extract the pdf, "pdftotext" 
or "pdftohtml", default "pdftotext".}

\item{converter.path}{Path to the pdftotext.exe or pdftohtml.exe executable
program. If NULL, make sure the xpdf folder is under R.home() or the parent 
directory of R.home(). E.g., for an x86 Win PC, \code{paste0(R.home(), 
paste0("/xpdfbin-win-4.00/bin32/pdftotext.exe"))}.}

\item{use.parallel}{Logical, whether to apply parallel computation when there 
are more than 50 files to process. Default \code{TRUE}.}
}
\value{
Full name of the PDFs if you set \code{silent=FALSE}.
}
\description{
Use Xpdf to convert PDFs to text files. The function relies on
\pkg{\link{doParallel}} to convert the PDFs in parallel.
It calls Xpdf to complete the conversion. Currently this function
only supports Windows.
}
\examples{
\dontrun{
convPDFs()
}
}
\author{
Yiying Wang, \email{wangy@aetna.com}
}
