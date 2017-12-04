#' @title List sheets in an xlsx (Excel) file
#'
#' @description
#' `xlsx_sheets()` returns the names of the sheets in a workbook, as a character
#' vector.  They are in the same order as they appear in the spreadsheet when it
#' is opened with a spreadsheet application like Excel or LibreOffice.
#'
#' @param path Path to the xlsx file.
#' @param check_filetype Logical. Whether to check that the filetype is xlsx (or
#' xlsm) by looking at the file itself, rather than using the filename
#' extension.
#'
#' @return
#' A character vector of the names of the worksheets in the file.
#'
#' @export
#' @examples
#' examples <- system.file("extdata/examples.xlsx", package = "tidyxl")
#' xlsx_sheet_names(examples)
xlsx_sheet_names <- function(path, check_filetype = TRUE) {
  utils_xlsx_sheet_files(path)$name
}
