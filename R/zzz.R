rpptx_py <- NULL
pptx <- NULL

# Make pptx accessible on package load
.onLoad <- function(libname, pkgname) {

  # Check if Python is installed and prompt the user to install if not
  python_is_installed <- reticulate::py_available()
  if (python_is_installed != TRUE) {
    wants_to_install_python <- menu(
      choices = c("Yes", "No"),
      title = "No Python installation was detected. Would you like to install one now?"
    )

    if (wants_to_install_python == 1) {
      reticulate::install_miniconda()
    } else {
      print("Cancelling installation")
    }
  }

  # Load the pptx module
  pptx <<- reticulate::import("pptx", delay_load = TRUE)

  # Load python code
  rpptx_py <<- reticulate::import_from_path(
    "rpptx_py",
    path = system.file("python", package = packageName())
  )
}
