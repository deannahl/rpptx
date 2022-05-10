pptx <- reticulate::import("pptx")

#' Create Presentation object from a PowerPoint file.
#'
#' @param path (character) File path for a PowerPoint file.
#'
#' @return Python pptx Presentation object.
#' @export
presentation <- function(path) {
  pptx$Presentation(path)
}

#' Save presentation as a new file
#'
#' @param pres (Presentation object) Python pptx Presentation object to save.
#' @param path (character) Path indicating where to save the presentation.
#'
#' @export
save_pres <- function(pres, path) {
  tryCatch(
    pres$save(path),
    error = function(e) {
      stop(paste0("Could not save the presentation to ", path,
                  ".\n  Check that the path is valid and the file is not open."))
    }
  )
}
