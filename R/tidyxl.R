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
#'
#' # All sheets
#' tidyxl(datasets)
#'
#' # Specific sheet either by position or by name
#' tidyxl(datasets, 2)
#' tidyxl(datasets, "mtcars")
tidyxl <- function(path, sheets = NA) {
  path <- check_file(path)
  all_sheets <- xlsx_sheets(path)
  if (anyNA(sheets)) { # This is no good -- what does it mean if vector length > 1?
    if (length(sheets) > 1) {
      warning("Argument 'sheets' included NAs, so every sheet in the workbook will be imported.")
    }
    sheets = all_sheets$name
  }
  sheets <- standardise_sheet(sheets, all_sheets)
  if (length(sheets) == 0) {
    warning("No sheets found.", call. = FALSE)
    return(list())
  }
  xlsx_read_(path, sheets)
}
