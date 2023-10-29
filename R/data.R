#' Names of all Excel functions
#'
#' A dataset containing the names of all functions available in Excel.  This is
#' useful for identifying user-defined functions in formulas tokenized by
#' [xlex()].
#'
#' Note that this includes future function names that are already reserved.
#'
#' @format A character vector of length 600.
#' @source Pages 26--27 of Microsoft's document "Excel (.xlsx) extensions to the
#' office openxml spreadsheetml file format p.24"
#' \url{https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e},
#' revision 8.0 2017-06-20.
"excel_functions"

#' Names and RGB values of Excel standard colours
#'
#' A dataset containing the names and RGB colour values of Excel's standard
#' palette.
#'
#' @format A data frame with 10 rows and 2 variables:
#'   * `name` Name of the colour
#'   * `rgb`  RGB value of the colour
"xlsx_color_standard"

#' @rdname xlsx_color_standard
"xlsx_colour_standard"
