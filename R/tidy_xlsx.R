#' @title Import xlsx (Excel) cell contents into a tidy structure.
#'
#' @description
#' \code{tidy_xlsx} imports data from spreadsheets without coercing it into a
#' rectangle.  Each cell is represented by a row in a data frame, giving the
#' cell's address, contents, formula, height, width, and keys to look up the
#' cell's formatting in an adjacent data structure within the list returned by
#' this function.
#'
#' @param path Path to the xlsx file.
#' @param sheets Sheets to read. Either a character vector (the names of the
#' sheets), an integer vector (the positions of the sheets), or NA (default, all
#' sheets).
#'
#' @details
#' A cell has two 'values': its content, and sometimes also a formula.  It also
#' has formatting applied at the 'style' level, which can be locally overriden.
#'
#' \subsection{content}{
#'   Depending on the cell, the content may be a numeric value such as 365 or
#'   365.25, it may represent a date/datetime in one of Excel's date/datetime
#'   systems, or it may be an index into an internal table of strings.
#'   \code{tidy_xlsx} attempts to infer the correct data type of each cell,
#'   returning its value in the appropriate column (error, logical, numeric,
#'   date, character). In case this cleverness is unhelpful, the unparsed value
#'   and type information is available in the 'content' and 'type' columns.
#' }
#'
#' \subsection{formula}{
#'   When a cell has a formula, the value in the 'content' column is the result
#'   of the formula the last time it was evaluated.
#'
#'   Certain groups of cells may share a formula that differs only by addresses
#'   refered to in the formula; such groups are identified by an index, the
#'   'formula_group'.  The xlsx (Excel) file format only records the formula
#'   against one cell in any group.  It is planned for \code{tidy_xlsx} to parse
#'   such formulas and copy them to the other cells in a group, making the
#'   necessary changes to addresses in the formula.
#'
#'   Array formulas may also apply to a group of cells, identified by an address
#'   'formula_ref', but xlsx (Excel) file format only records the formula
#'   against one cell in the group.  It is planned for \code{tidy_xlsx} to parse
#'   such addresses and copy the array formula to the other cells in the group.
#'   Unlike shared formulas, no changes to addresses in array formulas are
#'   necessary.
#'
#'   Formulas that refer to other workbooks currently do not name the workbooks
#'   directly, instead via indices such as \code{[1]}.  It is planned to
#'   dereference these.
#' }
#'
#' \subsection{formatting}{
#'   Cell formatting is returned in \code{x$formats}.  There are two types of
#'   formatting: 'style' formatting, such as Excel's built-in styles 'normal',
#'   'bad', etc., and 'local' formatting, which overrides the style.  These are
#'   returned in \code{x$formats$style} and \code{x$formats$local}, with
#'   identical structures.  To look up the local formatting of a given cell,
#'   take the cell's 'local_format_id' value (\code{x$data$Sheet1[1,
#'   "local_format_id"]}), and use it as an index into the format structure.
#'   E.g. to look up the font size,
#'   \code{x$formats$local$font$size[local_format_id]}.  To see all available
#'   formats, type `str(x$formats$local)`.
#' }
#'
#' @return
#' A list of the data within each sheet ($data), and the formatting applied to
#' each cell ($formats).
#'
#' Each sheet's data is returned as a data frames, one per sheet, by the sheet
#' name.  For example, the data in a sheet named 'My Worksheet' is in
#' x$data$`My Worksheet`.  Each data frame has the following
#' columns:
#'
#' \describe{
#'   \item{address}{The cell address in A1 notation.}
#'   \item{row}{The row number of a cell address (integer).}
#'   \item{col}{The column numer of a cell address (integer).}
#'   \item{content}{The content of a cell before type inference (see
#'   'Details').}
#'   \item{formula}{The formula in a cell (see 'Details').}
#'   \item{formula_type}{NA for ordinary formulas, or 'array' for array
#'   formulas.}
#'   \item{formula_ref}{The address (in A1 notation) of the cell that defines
#'   the formula of this cell (see 'Details').}
#'   \item{formula_group}{The formula group to which the cell belongs (see
#'   'Details').}
#'   \item{formula_ref}{The address of a range of cells group to which an array
#'   formula or shared formula applies (see 'Details').}
#'   \item{formula_group}{An index of a group of cells to which a shared formula
#'   applies (see 'Details').}
#'   \item{type}{The type of a cell in Excel's notation (b = boolean, e = error,
#'   s = string, str = formula).}
#'   \item{data_type}{The type of a cell, referring to the following columns
#'   (error, logical, numeric, date).}
#'   \item{error}{The error value of a cell.}
#'   \item{logical}{The boolean value of a cell.}
#'   \item{numeric}{The numeric value of a cell.}
#'   \item{date}{The date value of a cell.}
#'   \item{character}{The string value of a cell.}
#'   \item{comment}{The text of a comment attached to a cell.}
#'   \item{height}{The height of a cell's row, in Excel's units.}
#'   \item{width}{The width of a cell's column, in Excel's units.}
#'   \item{style_format}{An index into a table of style formats
#'   \code{x$formats$style} (see 'Details').}
#'   \item{local_format_id}{An index into a table of local cell formats
#'   \code{x$formats$local} (see 'Details').}
#' }
#'
#' Cell formatting is returned in \code{x$formats}.  There are two types or
#' scopes of formatting: 'style' formatting, such as Excel's built-in styles
#' 'normal', 'bad', etc., and 'local' formatting, which overrides particular
#' elements of the style, e.g. by making it bold.  Both types of  are returned
#' in \code{x$formats$style} and \code{x$formats$local}, with identical
#' structures.  To look up the local formatting of a given cell, take the cell's
#' 'local_format_id' value (\code{x$data$Sheet1[1, "local_format_id"]}), and use
#' it as an index into the format structure.  E.g. to look up the font size,
#' \code{x$formats$local$font$size[local_format_id]}.  To see all available
#' formats, type `str(x$formats$local)`.
#'
#' Colours may be recorded in any of three ways: a hexadecimal RGB string with
#' or without alpha, an 'indexed' colour, and an index into a 'theme'.
#' \code{tidy_xlsx} dereferences 'indexed' and 'theme' colours to their hexadecimal
#' RGB string representation, and standardises all RGB strings to have an alpha
#' channel in the first two characters.  The 'index' and the 'theme' name are
#' still provided.  To filter by an RGB string, you could  look up the RGB
#' values in a spreadsheet program (e.g. Excel, LibreOffice, Gnumeric), and use
#' the \code{\link{rgb}} function to convert these to a hexadecimal string.
#'
#' @export
#' @examples
#' examples <- system.file("extdata/examples.xlsx", package = "tidyxl")
#'
#' # All sheets
#' str(tidy_xlsx(examples)$data)
#'
#' # Specific sheet either by position or by name
#' str(tidy_xlsx(examples, 2)$data)
#' str(tidy_xlsx(examples, "Sheet1")$data)
#'
#' # Data (cell values)
#' x <- tidy_xlsx(examples)
#' str(x$data$Sheet1)
#'
#' # Formatting
#' str(x$formats$local)
#'
#' # The formats of particular cells can be retrieved like this:
#'
#' Sheet1 <- x$data$Sheet1
#' x$formats$style$font$bold[Sheet1$style_format]
#' x$formats$local$font$bold[Sheet1$local_format_id]
#'
#' # To filter for cells of a particular format, first filter the formats to get
#' # the relevant indices, and then filter the cells by those indices.
#' bold_indices <- which(x$formats$local$font$bold)
#' Sheet1[Sheet1$local_format_id %in% bold_indices, ]
tidy_xlsx <- function(path, sheets = NA) {
  path <- check_file(path)
  all_sheets <- xlsx_sheets(path)
  if (anyNA(sheets)) {
    if (length(sheets) > 1) {
      warning("Argument 'sheets' included NAs, which were discarded.")
      sheets = sheets[!is.na(sheets)]
      if (length(sheets) == 0) {
        stop("All elements of argument 'sheets' were discarded.")
      }
    } else {
      sheets = all_sheets$index
    }
  }
  sheets <- standardise_sheet(sheets, all_sheets)
  xlsx_read_(path, sheets$index, sheets$name, sheets$comments_path)
}
