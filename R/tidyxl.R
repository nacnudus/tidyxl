#' Read xlsx data and formatting
#'
#' @param path Path to the xlsx file
#' @param sheet Sheet to read. Either a string (the name of a sheet), or
#'   an integer (the position of the sheet). Defaults to the first sheet.
#' @name tidyxl
NULL

#' @describeIn tidyxl Read xlsx data and formatting semantically (not as a table)
#' @export
#' @examples
#' datasets <- system.file("extdata/datasets.xlsx", package = "readxl")
#' tidyxl(datasets)
#'
#' # Specific sheet either by position or by name
#' tidyxl(datasets, 2)
#' tidyxl(datasets, "mtcars")
tidyxl <- function(path, sheet = 1) {
  path <- check_file(path)
  sheet <- standardise_sheet(sheet, xlsx_sheet_names(path))

  xlsx_read_(path, sheet)
}
