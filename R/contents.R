#' @title Import xlsx (Excel) cell contents into a tidy structure.
#'
#' @description
#' \code{contents} imports data from spreadsheets without coercing it into a
#' rectangle.  Each cell is represented by a row in a data frame, giving the
#' cell's address, contents, formula, height, width, and keys to look up the
#' cell's formatting in a separate data structure (see \code{\link{formats}}).
#'
#' @param path Path to the xlsx file.
#' @param sheets Sheets to read. Either a character vector (the names of the
#' sheets), an integer vector (the positions of the sheets), or NA (default, all
#' sheets).
#'
#' @details
#' A cell has two 'values': its content, and sometimes also a formula.
#'
#' \subsection{content}{
#'   Depending on the cell, the content may be a value such as 365 or 365.25, it
#'   may represent a date/datetime in one of Excel's date/datetime systems, or
#'   it may be an index into an internal table of strings.  In the last case
#'   (index into strings table), \code{contents} automatically looks up the
#'   string and returns it in the 'character' column.
#' }
#'
#' \subsection{formula}{
#'   When a cell has a formula, the value in the 'content' column is the result
#'   of the formula the last time it was evaluated.
#'
#'   Certain groups of cells may share a formula that differs only by addresses
#'   refered to in the formula; such groups are identified by an index, the
#'   'formula_group'.  The xlsx (Excel) file format only records the formula
#'   against one cell in any group.  It is planned for \code{contents} to parse
#'   such formulas and copy them to the other cells in a group, making the
#'   necessary changes to addresses in the formula.
#'
#'   Array formulas may also apply to a group of cells, identified by an address
#'   'formula_ref', but xlsx (Excel) file format only records the formula
#'   against one cell in the group.  It is planned for \code{contents} to parse
#'   such addresses and copy the array formula to the other cells in the group.
#'   Unlike shared formulas, no changes to addresses in array formulas are
#'   necessary.
#'
#'   Formulas that refer to other workbooks currently do not name the workbooks
#'   directly, instead via indices such as \code{[1]}.  It is planned to
#'   dereference these.
#' }
#'
#' @return
#' A list of data frames, one per sheet.  Each data frame has the following
#' columns:
#'
#' \describe{
#'   \item{address}{The cell address in A1 notation.}
#'   \item{row}{The row number of the cell address (integer).}
#'   \item{col}{The column numer of the cell address (integer).}
#'   \item{content}{The content of the cell (see 'Details').}
#'   \item{formula}{The formula in the cell (see 'Details').}
#'   \item{formula_group}{The formula group to which the cell belongs (see
#'   'Details').}
#'   \item{formula_ref}{The address of a range of cells group to which an array
#'   formula or shared formula applies (see 'Details').}
#'   \item{formula_group}{An index of a group of cells to which a shared formula
#'   applies (see 'Details').}
#'   \item{type}{The type of a cell (b = boolean, e = error, s = string, str =
#'   formula).}
#'   \item{character}{The string contents of a cell (see 'Details').}
#'   \item{height}{The height of the cell's row, in Excel's weird units.}
#'   \item{width}{The width of the cell's column, in Excel's weird units.}
#'   \item{style_format_id}{An index into a table of style formats (see
#'   \code{\link{formats}}).}
#'   \item{local_format_id}{An index into a table of local cell formats (see
#'   \code{\link{formats}}).}
#' }
#'
#' @export
#' @examples
#' examples <- system.file("extdata/examples.xlsx", package = "tidyxl")
#'
#' # All sheets
#' contents(examples)
#'
#' # Specific sheet either by position or by name
#' contents(examples, 2)
#' contents(examples, "Sheet1")
#'
#' # Note that single sheets are still returned in a list for consistency.
#' # Extract particular sheets using [[.
#' contents(examples)[[2]]
contents <- function(path, sheets = NA) {
  path <- check_file(path)
  all_sheets <- xlsx_sheets_(path)
  if (anyNA(sheets)) {
    if (length(sheets) > 1) {
      warning("Argument 'sheets' included NAs, which were discarded.")
      sheets = sheets[!is.na(sheets)]
    } else {
      sheets = all_sheets$index
    }
  }
  sheets <- standardise_sheet(sheets, all_sheets)
  if (nrow(sheets) == 0) {
    warning("No sheets found.", call. = FALSE)
    return(list())
  }
  imported <- xlsx_read_(path, sheets$index, sheets$name)
  sheet_data <- imported[["data"]]
  sheet_comments <- lapply(sheets$comments_path, comments, path = path)
  imported[["data"]] <-
    purrr::map2(sheet_data, sheet_comments,
                ~ dplyr::left_join(.x, .y, by = c("address" = "ref")))
  imported
}

# Parse comments (currently uses R, might use C++ in future).
# Comments are then joined to cell data by cell address (ref).
comments <- function(comments_path, path) {
  if (is.na(comments_path)) {
    return(tibble::tibble(ref = character(), comment = character()))
  }
  unz(path, comments_path) %>%
  xml2::read_xml() %>%
  xml2::xml_ns_strip() %>%
  xml2::xml_find_all("commentList") %>%
  xml2::xml_find_all("comment") %>%
  {tibble::tibble(ref = xml2::xml_attr(., "ref"), comment = xml2::xml_text(.))}
}
