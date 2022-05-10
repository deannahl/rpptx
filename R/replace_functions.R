
#' Replace an image in a PowerPoint text box.
#'
#' @param pres Python pptx Presentation object.
#' @param label (character) Label of the image to replace.
#' @param new_image (character) New image to replace the old image.
#' @param new_height (boolean) Should the aspect ratio of the new image be preserved (`TRUE`) or
#' should the image height be set to match the height of the old image (`FALSE`).
#'
#' @export
replace_image <- function(pres, label, new_image, new_height=TRUE) {
  rpptx_py$py_replace_image(pres, label, new_image, new_height)
}

#' Replace text in a PowerPoint table while retaining formatting.
#'
#' @param pres Python pptx Presentation object.
#' @param label (character) Label of the text box containing the table to replace
#' @param new_table (data frame) Data frame to replace elements in the old table
#' @param col_names (boolean) Keep new_table column names as the first row of the
#'  new table?
#'
#' @return
#' @export
replace_table <- function(pres, label, new_table, col_names = TRUE) {
  # If col_names is TRUE, insert a new row containing the column names
  if (col_names == TRUE) {
    new_table <- rbind(colnames(new_table), new_table)
  }

  # If any values are NA, replace with ""
  new_table[is.na(new_table)] <- ""

  # Get the shape of the new table to check it against the old table
  new_table_dim <- c(nrow(new_table), ncol(new_table))

  # Convert new_table to a list (by row)
  new_table <- as.list(t(new_table))

  rpptx_py$py_replace_table(pres, label, new_table, new_table_dim)
}

#' Replace text in a PowerPoint text box while retaining formatting.
#'
#' @param pres Python pptx Presentation object.
#' @param label (character) Label of the text box containing the text to replace.
#' @param new_text (character) New text to replace the old text.
#'
#' @return
#' @export
replace_text <- function(pres, label, new_text) {
  rpptx_py$py_replace_text(pres, label, new_text)
}

#' Replace category plot
#'
#' @param pres Python pptx Presentation object.
#' @param label (character) Label of the text box containing the text to replace.
#' @param data (data frame) Data frame used to replace the old table.
#'
#' @return
#' @export
replace_category_plot <- function(pres, label, data) {
  categories <- as.character(unlist(data[,1]))
  series_levels <- colnames(data)[-1]
  series_values <- unname(as.list(data[,-1]))
  rpptx_py$py_replace_category_plot(pres, label, categories, series_levels, series_values)
}

#' Replace a donut plot
#'
#' @param pres Python pptx Presentation object.
#' @param label (character) Label of the text box containing the text to replace.
#' @param percent_favorable (double) Proportion of favorable responses.
#'
#' @return
#' @export
replace_donut_plot <- function(pres, label, percent_favorable) {
  rpptx_py$py_replace_category_plot(pres, label,
                                    categories=c('Favorable', 'Unfavorable'),
                                    series_levels=c('Series1'),
                                    series_values=list(c(percent_favorable, 1-percent_favorable)))
}

#' Get the unique slide ID for a slide using the slide number
#'
#' @param pres  Python pptx Presentation object.
#' @param slide_num (int) The slide number of the target slide.
#'
#' @return
#' @export
get_slide_id <- function(pres, slide_num) {
  rpptx_py$get_slide_id(pres, slide_num)
}
