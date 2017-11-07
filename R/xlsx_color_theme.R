#' @title Import theme color definitions from xlsx (Excel) files
#'
#' @description
#' `xlsx_color_theme()` (alias `xlsx_colour_theme()` returns the names and RGB
#' values of theme colours defined in xlsx (Excel) files.  For example,
#' `"accent6"` is the name of a theme colour in Excel, which could resolve to
#' any RGB colour defined by the author of the file.  Themes are often defined
#' to comply with corporate standards.
#'
#' @param path Path to the xlsx file.
#'
#' @return
#' A data frame, one row per colour, with the following columns.
#'
#' * `theme` The name of the theme.
#' * `name`
#' * `rgb` The RGB colour that has been set for the theme in this file.
#'
#' @export
#' @examples
#' examples <- system.file("extdata/examples.xlsx", package = "tidyxl")
#' xlsx_color_theme(examples)
#' xlsx_colour_theme(examples)
xlsx_color_theme <- function(path) {
  path <- check_file(path)
  xlsx_color_theme_(path)
}

#' @rdname xlsx_color_theme
#' @export
xlsx_colour_theme <- xlsx_color_theme
