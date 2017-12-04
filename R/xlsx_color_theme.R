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
#' @param check_filetype Logical. Whether to check that the filetype is xlsx (or
#' xlsm) by looking at the file itself, rather than using the filename
#' extension.
#'
#' @return
#' A data frame, one row per colour, with the following columns.
#'
#' * `name` The name of the theme.
#' * `rgb` The RGB colour that has been set for the theme in this file.
#'
#' @export
#' @examples
#' examples <- system.file("extdata/examples.xlsx", package = "tidyxl")
#' xlsx_color_theme(examples)
#' xlsx_colour_theme(examples)
xlsx_color_theme <- function(path, check_filetype = TRUE) {
  path <- check_file(path)
  xlsx_color_theme_(path)
}

#' @rdname xlsx_color_theme
#' @export
xlsx_colour_theme <- xlsx_color_theme
