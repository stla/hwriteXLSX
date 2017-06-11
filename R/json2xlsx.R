#' @useDynLib json2xlsx
.json2xlsx <- function(json, outfile){
  invisible(.C("json2xlsx", json=json, outfile=outfile))
}

#' Write a XLSX file from a JSON string
#'
#' @param json JSON string
#' @param outfile name of output file
#' @param overwrite logical, whether to overwrite `outfile` if it exists
#'
#' @return No returned value.
#' @export
#' @importFrom jsonlite validate
#'
#' @examples
#' colorMatrix <- matrix(grDevices::colours(), ncol=9)
#' f <- function(i, j){
#'   cell(j, i, value = colorMatrix[i,j], color = colorMatrix[i,j])
#' }
#' sheet <- list(Colors = do.call(Vectorize(f), expand.grid(i=1:73, j=1:9)))
#' json <- jsonlite::toJSON(sheet, null="null", auto_unbox = TRUE)
#' \donttest{json2xlsx(json, "colors.xlsx")}
json2xlsx <- function(json, outfile, overwrite=FALSE){
  if(!jsonlite::validate(json)){
    stop("Invalid JSON.")
  }
  if(!overwrite && file.exists(outfile)){
    stop(sprintf("File `%s` found.", outfile))
  }
  .json2xlsx(json, outfile)
}
