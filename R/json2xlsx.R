#' @useDynLib json2xlsx
.json2xlsx <- function(json1, json2, json3, outfile){
  invisible(.C("json2xlsx", json1=json1, json2=json2, json3=json3, outfile=outfile))
}

#' @importFrom readr write_file
json2xlsx2_exe <- function(json1, json2="{}", json3="{}", outfile){
  jsonFile1 <- tempfile(fileext = ".json")
  readr::write_file(as.character(json1), path=jsonFile1)
  jsonFile2 <- tempfile(fileext = ".json")
  readr::write_file(as.character(json2), path=jsonFile2)
  jsonFile3 <- tempfile(fileext = ".json")
  readr::write_file(as.character(json3), path=jsonFile3)
  if(.Platform$OS.type == "windows"){
    exe <- system.file("bin", "windows", "writexlsx.exe", package="hwriteXLSX")
  }else{
    exe <- system.file("bin", "unix", "writexlsx", package="hwriteXLSX")
  }
  command <- sprintf("%s -c %s -i %s -p %s -o %s", exe,
                     shQuote(jsonFile1), shQuote(jsonFile2),
                     shQuote(jsonFile3), shQuote(outfile))
  system(command)
}

#' Write a XLSX file from a JSON string
#'
#' @param json1 JSON string for the cells
#' @param json2 JSON string for the images
#' @param json3 JSON string for the passwords
#' @param outfile name of output file
#' @param overwrite logical, whether to overwrite `outfile` if it exists
#'
#' @return No returned value.
#' @export
#' @importFrom jsonlite validate
#'
#' @note This function is rather for internal purpose. The user should use
#' \code{\link{hwriteXLSX}}.
#'
#' @examples
#' colorMatrix <- matrix(grDevices::colours(), ncol=9)
#' f <- function(i, j){
#'   cell(j, i, value = colorMatrix[i,j], color = colorMatrix[i,j])
#' }
#' sheet <- list(Colors = do.call(Vectorize(f), expand.grid(i=1:73, j=1:9)))
#' json <- jsonlite::toJSON(sheet, null="null", auto_unbox = TRUE)
#' \donttest{json2xlsx(json, outfile="colors.xlsx")}
json2xlsx <- function(json1, json2="{}", json3="{}", outfile, overwrite=FALSE, dll=FALSE){
  if(!overwrite && file.exists(outfile)){
    stop(sprintf("File `%s` already exists.", outfile))
  }
  if(!jsonlite::validate(json1)){
    stop("Invalid JSON string `json1`.")
  }
  if(!jsonlite::validate(json2)){
    stop("Invalid JSON string `json2`.")
  }
  if(!jsonlite::validate(json3)){
    stop("Invalid JSON string `json3`.")
  }
  if(json1 == "{}" && json2 == "{}"){
    stop("`json1` and `json2` cannot be both empty.")
  }
  if(dll){
    .json2xlsx(json1, json2, json3, outfile)
  }else{
    json2xlsx2_exe(json1, json2, json3, outfile)
  }
}


