#' @importFrom stringr str_detect
searchXpdfFolder <- function(){
    r.parent <- dirname(R.home())
    
    dirs <- c(list.dirs(r.parent, recursive=FALSE),
              list.dirs(R.home(), recursive=FALSE))
    if (Sys.getenv("xpdfdir") != ""){
        xpdfdir <- Sys.getenv("xpdfdir")
    }else if (any(str_detect(dirs, "xpdfbin"))){
        xpdfdir <- dirs[str_detect(tolower(dirs), "xpdfbin")][1]
    }else{
        return(NULL)
    }
    return(paste(c(
        unlist(strsplit(xpdfdir, "[/\\\\]")), 
        paste0("bin", getOption("mach.arch"))), collapse="/"))
}

