#' @importFrom grDevices colours
createFont <- function(name=NULL, bold=NULL, color=NULL){
  if(!is.null(bold) && !is.logical(bold)){
    stop("`bold` must be logical.")
  }
  if(!is.null(color)){
    if(! color %in% grDevices::colours()){
      stop("Invalid color name.")
    }
  }
  if(!is.null(name) && !is.character(name)){
    stop("`name` must be a font name")
  }
  list(name=name, bold=bold, color=color)
}

createFormat <- function(numberFormat=NULL, font=NULL, fill=NULL){
  if(!is.null(fill)){
    if(! fill %in% grDevices::colours()){
      stop("Invalid color name.")
    }
  }
  list(numberFormat=numberFormat, font=font, fill=fill)
}

createComment <- function(text=NULL, author=NULL){
  if(is.null(text))
    return(NULL)
  else
    return(list(text=text, author=author))
}

createCell <- function(value=NULL, format=NULL, comment=NULL,
                       colspan=NULL, rowspan=NULL){
  return(list(value=value, format=format, comment=comment,
              colspan=colspan, rowspan=rowspan))
}

int_to_letter <- function(col) {
  #stopifnot(is.numeric(y))
  ret <- integer()
  while (col > 0) {
    remainder <- ((col - 1)%%26) + 1
    ret <- c(remainder, ret)
    col <- (col - remainder)%/%26
  }
  paste0(LETTERS[ret], collapse = "")
}
