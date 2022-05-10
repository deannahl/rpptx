rpptx_py <- NULL
pptx <- NULL

# Make pptx accessible on package load
.onLoad <- function(libname, pkgname) {

  # Load the pptx module
  pptx <<- reticulate::import("pptx", delay_load = TRUE)

  # Load python code
  rpptx_py <<- reticulate::import_from_path(
    "rpptx_py",
    path = system.file("python", package = packageName())
  )
}
