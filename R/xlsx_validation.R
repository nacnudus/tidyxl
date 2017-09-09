#' @title Import data validation rules of cells in xlsx (Excel) files
#'
#' @description
#' `xlsx_validation()` returns the data validation rules applied to cells in
#' xlsx (Excel) files.  Data validation rules control what constants can be
#' entered into a cell, e.g. any whole number between 0 and 9, or one of several
#' values from another part of the spreadsheet.
#'
#' Each data validation rule can only apply to cells on one worksheet, so
#' `xlsx_validation()` returns a named list of data frames, one for each
#' worksheet, describing the rules on that worksheet.
#'
#' @param path Path to the xlsx file.
#' @param sheets Sheets to read. Either a character vector (the names of the
#' sheets), an integer vector (the positions of the sheets), or NA (default, all
#' sheets).
#'
#' @return
#' A named list of data frames, one for each worksheet in the file.  Each data
#' frame has the following columns.
#'
#' * `ref` Comma-delimited cell addresses to which the rules apply,
#'     e.g. `A106` or A115,A121:A122`.
#' * `type Data type of input, one of `whole`, `decimal`, `list`, `date`,
#'      `time`, `textLength`, `custom`, and `whole`.
#' * `operator` Unless `type` is `list` or `custom`, then `operator` is one of
#'     `between`, `notBetween`, `equal`, `notEqual`, `greaterThan`, `lessThan`,
#'     `greaterThanOrEqual`, `lessthanOrEqual`.
#' * `formula1` If `type` is `list`, then a range of cells whose values are
#'     allowed by the rule.  If `type` is `custom`, then a formula to determine
#'     alowable values.  Otherwise, a cell address or constant, coerced to
#'     character.  Dates and times are formatted like "2017-01-27 13:30:45".
#'     Times without dates are formatted like "13:30:45".
#' * `formula2` If `operator` is `between` or `notBetween`, then a cell address
#'     or constant as with formula1, otherwise NA.
#' * `allow_blank` Boolean, whether or not the rule allows blanks.
#' * `show_input_message` Boolean, whether or not the rule shows a message when
#'     the user begins entering a value.
#' * `prompt_title` Text to appear in the title bar of a popup message box
#'     when the user begins entering a value.
#' * `prompt_body` Text to appear in a popup message box when the user begins
#'     entering a value.  When `NA`, then some default text is shown.
#' * `show_error_message` Boolean, whether or not the rule shows a message when
#'     the user has entered a forbidden value.  When `NA`, then some default
#'     text is shown.
#' * `error_title` Text to appear in the title bar of a popup message box
#'     when the user enters a forbidden value.  When `NA`, then some default
#'     text is shown.
#' * `error_body` Text to appear in a popup message box when the user enters a
#'     forbidden value.  When `NA`, then some default text is shown.
#' * `error_symbol` Name of a symbol to appear in the popup error message when
#'     the user enters a forbidden value.
#'
#' @export
#' @examples
#' examples <- system.file("extdata/examples.xlsx", package = "tidyxl")
#' xlsx_validation(examples)
#' xlsx_validation(examples, 1)[["Sheet1"]]
#' xlsx_validation(examples, "Sheet1")[[1]]
xlsx_validation <- function(path, sheets = NA) {
  path <- check_file(path)
  all_sheets <- utils_xlsx_sheet_files(path)
  if (anyNA(sheets)) {
    if (length(sheets) > 1) {
      warning("Argument 'sheets' included NAs, which were discarded.")
      sheets <- sheets[!is.na(sheets)]
      if (length(sheets) == 0) {
        stop("All elements of argument 'sheets' were discarded.")
      }
    } else {
      sheets <- all_sheets$order
    }
  }
  sheets <- standardise_sheet(sheets, all_sheets)
  out <- xlsx_validation_(path, sheets$sheet_path, sheets$name)
  out <- lapply(out, function(x) x[order(x$ref), ])
  out
}
