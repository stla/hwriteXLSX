.onLoad <- function(libname, pkgname) {
  library.dynam("json2xlsx", pkgname, libname)
  .C("HsStart")
  invisible()
}

.onUnLoad <- function(libpath) {
  library.dynam.unload("json2xlsx", libpath)
  invisible()
}

