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

createFormat <- function(numberFormat=NULL, font=NULL){
  list(numberFormat=numberFormat, font=font)
}

createCell <- function(value=NULL, format=NULL, comment=NULL){
  return(list(value=value, format=format, comment=comment))
}
