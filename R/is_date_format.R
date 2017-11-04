#' Test that Excel number formats are date formats
#'
#' @description `is_date_format()` tests whether an Excel number format string
#' would a resolve to a date in Excel.  For example, the number format string
#' `"yyyy-mm-dd"` would resolve to a date, whereas the string `"0.0\\%"` would
#' not.
#'
#' This is useful if a cell formula contains a number formatting string (e.g.
#' `TEXT(45678,"yyyy")`), and you need to know that the constant 45678 is a date
#' in order to recover it at full resolution (rather than parsing the character
#' output "2025" as a year).
#'
#' It is used internally to convert the value of a cell to the correct data
#' type.
#'
#' @param x character vector of number format strings
#' @export
#' @examples
#' is_date_format(c("yyyy-mm-dd", "0.0%", "h:m:s", "£#,##0;[Red]-£#,##0"))
is_date_format <- function(x) {
  is_date_format_(x)
}
