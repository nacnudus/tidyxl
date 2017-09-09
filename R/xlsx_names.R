#' @title Import named formulas from xlsx (Excel) files
#'
#' @description
#' `xlsx_names()` returns the names and definitions of named formulas (aka named
#' ranges) in xlsx (Excel) files.
#'
#' Most names refer to ranges of cells, but they can also be defined as
#' formulas.  `xlsx_names()` tells you whether or not they are a range, using
#' [tidyxl::is_range()] to work this out.
#'
#' Names are scoped either globally (used only once in the file), or locally to
#' each sheet (can be reused with different definitions in diffent sheets).  For
#' sheet-scoped names, `xlsx_names()` provides the name of the sheet.
#'
#' @param path Path to the xlsx file.
#'
#' @return
#' A data frame, one row per name, with the following columns.
#'
#' * `name`
#' * `formula` Usually a range of cells, but sometimes a whole formula, e.g.
#'     `MAX(A2,1)`.
#' * `sheet_name` If the name is defined only for a specific sheet, the name of
#'     the sheet.  Otherwise `NA` for names defined globally.
#' * `is_range` Whether or not the `formula` is a range of cells.  This is handy
#'     for joining to the set of cells referred to by a name.  In this context,
#'     commas between cell addresses are always regarded as union operators --
#'     this differs from [tidyxl::xlex()], see that help file for details.
#'
#' @export
#' @examples
#' examples <- system.file("extdata/examples.xlsx", package = "tidyxl")
#' xlsx_names(examples)
xlsx_names <- function(path) {
  path <- check_file(path)
  out <- xlsx_names_(path)
  # Microsoft docs don't say what sheet_id links to.  Testing suggests it is the
  # 'id' of
  # except that it isn't the 'id' of utils_xlsx_sheet_files().  For now, add 1
  # and interpret is as the 'order' of utils_xlsx_sheet_files().
  sheets <- utils_xlsx_sheet_files(path)[, c("name", "id")]
  names(sheets)[1] <- "sheet_name"
  out$sheet_id <- out$sheet_id + 1
  out <- merge(out, sheets, by.x = "sheet_id", by.y = "id", all.x = TRUE)
  out$sheet_id <- NULL
  out$is_range <- is_range(out$formula)
  out
}
