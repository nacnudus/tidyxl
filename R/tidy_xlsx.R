#' @title Import xlsx (Excel) cell contents into a tidy structure.
#'
#' @description
#' `tidy_xlsx()` is deprecated.  Please use [xlsx_cells()] or [xlsx_formats()]
#' instead.
#'
#' `tidy_xlsx()` imports data from spreadsheets without coercing it into a
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
#' has formatting applied at the 'style' level, which can be locally overridden.
#'
#' \subsection{Content}{
#'   Depending on the cell, the content may be a numeric value such as 365 or
#'   365.25, it may represent a date/datetime in one of Excel's date/datetime
#'   systems, or it may be an index into an internal table of strings.
#'   `tidy_xlsx()` attempts to infer the correct data type of each cell,
#'   returning its value in the appropriate column (error, logical, numeric,
#'   date, character). In case this cleverness is unhelpful, the unparsed value
#'   and type information is available in the 'content' and 'type' columns.
#' }
#'
#' \subsection{Formula}{
#'   When a cell has a formula, the value in the 'content' column is the result
#'   of the formula the last time it was evaluated.
#'
#'   Certain groups of cells may share a formula that differs only by addresses
#'   referred to in the formula; such groups are identified by an index, the
#'   'formula_group'.  The xlsx (Excel) file format only records the formula
#'   against one cell in any group, but `tidy_xlsx()` propagates the formula to
#'   all the cells in the group, making the necessary changes to relative
#'   addresses in the formula.
#'
#'   Array formulas may also apply to a group of cells, identified by an address
#'   'formula_ref', but xlsx (Excel) file format only records the formula
#'   against one cell in the group.  Unlike ordinary formulas, `tidy_xlsx()`
#'   does not propagate these to the other cells in the group.
#'
#'   Formulas that refer to other workbooks currently do not name the workbooks
#'   directly, instead via indices such as `[1]`.  It is planned to
#'   dereference these.
#' }
#'
#' \subsection{Formatting}{
#'   Cell formatting is returned in `x$formats`.  There are two types of
#'   formatting: 'style' formatting, such as Excel's built-in styles 'normal',
#'   'bad', etc., and 'local' formatting, which overrides the style.  These are
#'   returned in `x$formats$style` and `x$formats$local`, with
#'   identical structures.  To look up the local formatting of a given cell,
#'   take the cell's 'local_format_id' value (`x$data$Sheet1[1,
#'   "local_format_id"]`), and use it as an index into the format structure.
#'   E.g. to look up the font size,
#'   `x$formats$local$font$size[local_format_id]`.  To see all available
#'   formats, type `str(x$formats$local)`.
#' }
#'
#' @return
#' A list of the data within each sheet (`$data`), and the formatting applied to
#' each cell (`$formats`).
#'
#' Each sheet's data is returned as a data frames, one per sheet, by the sheet
#' name.  For example, the data in a sheet named 'My Worksheet' is in
#' x$data$`My Worksheet`.  Each data frame has the following
#' columns:
#'
#' * `address` The cell address in A1 notation.
#' * `row` The row number of a cell address (integer).
#' * `col` The column number of a cell address (integer).
#' * `content` The content of a cell before type inference (see
#'     'Details').
#' * `formula` The formula in a cell (see 'Details').
#' * `formula_type` NA for ordinary formulas, or 'array' for array
#'     formulas.
#' * `formula_ref` The address (in A1 notation) of the cell that defines
#'     the formula of this cell (see 'Details').
#' * `formula_group` The formula group to which the cell belongs (see
#'     'Details').
#' * `formula_ref` The address of a range of cells group to which an array
#'     formula or shared formula applies (see 'Details').
#' * `formula_group` An index of a group of cells to which a shared formula
#'     applies (see 'Details').
#' * `type` The type of a cell in Excel's notation (b = boolean, e = error,
#'     s = string, str = formula).
#' * `data_type` The type of a cell, referring to the following columns
#'     (error, logical, numeric, date).
#' * `error` The error value of a cell.
#' * `logical` The boolean value of a cell.
#' * `numeric` The numeric value of a cell.
#' * `date` The date value of a cell.
#' * `character` The string value of a cell.
#' * `comment` The text of a comment attached to a cell.
#' * `height` The height of a cell's row, in Excel's units.
#' * `width` The width of a cell's column, in Excel's units.
#' * `style_format` An index into a table of style formats
#'     `x$formats$style` (see 'Details').
#' * `local_format_id` An index into a table of local cell formats
#'     `x$formats$local` (see 'Details').
#'
#' Cell formatting is returned in `x$formats`.  There are two types or
#' scopes of formatting: 'style' formatting, such as Excel's built-in styles
#' 'normal', 'bad', etc., and 'local' formatting, which overrides particular
#' elements of the style, e.g. by making it bold.  Both types of  are returned
#' in `x$formats$style` and `x$formats$local`, with identical
#' structures.  To look up the local formatting of a given cell, take the cell's
#' 'local_format_id' value (`x$data$Sheet1[1, "local_format_id"]`), and use
#' it as an index into the format structure.  E.g. to look up the font size,
#' `x$formats$local$font$size[local_format_id]`.  To see all available
#' formats, type `str(x$formats$local)`.
#'
#' Colours may be recorded in any of three ways: a hexadecimal RGB string with
#' or without alpha, an 'indexed' colour, and an index into a 'theme'.
#' `tidy_xlsx` dereferences 'indexed' and 'theme' colours to their hexadecimal
#' RGB string representation, and standardises all RGB strings to have an alpha
#' channel in the first two characters.  The 'index' and the 'theme' name are
#' still provided.  To filter by an RGB string, you could  look up the RGB
#' values in a spreadsheet program (e.g. Excel, LibreOffice, Gnumeric), and use
#' the [grDevices::rgb()] function to convert these to a hexadecimal string.
#'
#' @export
#' @examples
#' \dontrun{
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
#' }
tidy_xlsx <- function(path, sheets = NA) {
  .Deprecated(msg = paste("'tidy_xlsx()' is deprecated.",
                          "Use 'xlsx_cells()' or 'xlsx_formats()' instead.",
                          sep = "\n"))
  path <- check_file(path)
  all_sheets <- utils_xlsx_sheet_files(path)
  sheets <- check_sheets(sheets, path)
  formats <- xlsx_formats_(path)
  cells <- xlsx_cells_(path, sheets$sheet_path, sheets$name, sheets$comments_path)
  # Split into a list of data frames, one per sheet
  cells$sheet <- factor(cells$sheet, levels = sheets$name) # control sheet order
  cells_list <- split(cells, cells$sheet)
  cells_list <- lapply(cells_list, function(x) x[, -1])    # remove 'sheet' col
  list(data = cells_list, formats = formats)
}
