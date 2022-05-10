#' Install the Python pptx package
#'
#' @param method Installation method. By default, "auto" automatically finds a method that will work
#'   in the local environment. Change the default to force a specific installation method. Note that
#'   the "virtualenv" method is not available on Windows.
#' @param conda The path to a conda executable. Use "auto" to allow reticulate to automatically find
#'   an appropriate conda binary. See Finding Conda for more details.
#'
#' @export
install_pptx <- function(method = "auto", conda = "auto") {
  reticulate::py_install("pptx", method = method, conda = conda)
}
