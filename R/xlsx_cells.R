#' @title Import xlsx (Excel) cell contents into a tidy structure.
#'
#' @description
#' `xlsx_cells()` imports data from spreadsheets without coercing it into a
#' rectangle.  Each cell is represented by a row in a data frame, giving the
#' cell's address, contents, formula, height, width, and keys to look up the
#' cell's formatting in the return value of [tidyxl::xlsx_formats()].
#'
#' @param path Path to the xlsx file.
#' @param sheets Sheets to read. Either a character vector (the names of the
#' sheets), an integer vector (the positions of the sheets), or NA (default, all
#' sheets).
#' @param check_filetype Logical. Whether to check that the filetype is xlsx (or
#' xlsm) by looking at the file itself, rather than using the filename
#' extension.
#'
#' @return
#' A data frame with the following columns.
#'
#' * `sheet` The worksheet that the cell is from.
#' * `address` The cell address in A1 notation.
#' * `row` The row number of a cell address (integer).
#' * `col` The column number of a cell address (integer).
#' * `is_blank` Whether or not the cell has a value
#' * `data_type` The type of a cell, referring to the following columns:
#'     error, logical, numeric, date, character, blank.
#' * `error` The error value of a cell.
#' * `logical` The boolean value of a cell.
#' * `numeric` The numeric value of a cell.
#' * `date` The date value of a cell.
#' * `character` The string value of a cell.
#' * `formula` The formula in a cell (see 'Details').
#' * `is_array` Whether or not the formula is an array formula.
#' * `formula_ref` The address of a range of cells group to which an array
#'     formula or shared formula applies (see 'Details').
#' * `formula_group` The formula group to which the cell belongs (see
#'     'Details').
#' * `comment` The text of a comment attached to a cell.
#' * `height` The height of a cell's row, in Excel's units.
#' * `width` The width of a cell's column, in Excel's units.
#' * `style_format` An index into a table of style formats
#'     `x$formats$style` (see 'Details').
#' * `local_format_id` An index into a table of local cell formats
#'     `x$formats$local` (see 'Details').
#'
#' Cell formatting is returned in [tidyxl::xlsx_formats()].  There are two types
#' or scopes of formatting: 'style' formatting, such as Excel's built-in styles
#' 'normal', 'bad', etc., and 'local' formatting, which overrides particular
#' elements of the style, e.g. by making it bold.  Both types are returned, in
#' the `$style` and `$local` sublists of [tidyxl::xlsx_formats()], with
#' identical structures.  To look up the local formatting of a given cell, take
#' the cell's 'local_format_id' value (`my_cells$data$Sheet1[1,
#' "local_format_id"]`), and use it as an index into the format structure.  E.g.
#' to look up the font size, `my_formats$local$font$size[local_format_id]`.  To
#' see all available formats, type `str(my_formats$local)`.
#'
#' @details
#' A cell has two 'values': its content, and sometimes also a formula.  It also
#' has formatting applied at the 'style' level, which can be locally overridden.
#'
#' \subsection{Content}{
#'   Depending on the cell, the content may be a numeric value such as 365 or
#'   365.25, it may represent a date/datetime in one of Excel's date/datetime
#'   systems, or it may be an index into an internal table of strings.
#'   `xlsx_cells()` attempts to infer the correct data type of each cell,
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
#'   against one cell in any group.  `xlsx_cells()` propagates such formulas to
#'   the other cells in a group, making the necessary changes to relative
#'   addresses in the formula.
#'
#'   Array formulas may also apply to a group of cells, identified by an address
#'   'formula_ref', but xlsx (Excel) file format only records the formula
#'   against one cell in the group.  `xlsx_cells()` propagates such formulas to
#'   the other cells in a group.  Unlike shared formulas, no changes to
#'   addresses in array formulas are necessary.
#'
#'   Formulas that refer to other workbooks currently do not name the workbooks
#'   directly, instead via indices such as `[1]`.  It is planned to
#'   dereference these.
#' }
#'
#' \subsection{Formatting}{
#'   Cell formatting is returned by [tidyxl::xlsx_formats()].  There are two
#'   types of formatting: 'style' formatting, such as Excel's built-in styles
#'   'normal', 'bad', etc., and 'local' formatting, which overrides the style.
#'   These are returned in the `$style` and `$local` sublists of
#'   [tidyxl::xlsx_formats()], with identical structures.
#'
#'   To look up the local formatting of a given cell, take the cell's
#'   `local_format_id` value (`my_cells$Sheet1[1, "local_format_id"]`), and use
#'   it as an index into the format structure.  E.g. to look up the font size,
#'   `my_formats$local$font$size[local_format_id]`.  To see all available
#'   formats, type `str(my_formats$local)`.
#'
#'   Strings can be formatted within a cell, so that a single cell can contain
#'   substrings with different formatting.  This in-cell formatting is available
#'   in the column `character_formatted`, which is a list-column of data frames.
#'   Each row of each data frame describes a substring and its formatting.  For
#'   cells without a character value, `character_formatted` is `NULL`, so for
#'   further processing you might need to filter out the `NULL`s first.
#' }
#'
#' @export
#' @examples
#' examples <- system.file("extdata/examples.xlsx", package = "tidyxl")
#'
#' # All sheets
#' str(xlsx_cells(examples))
#'
#' # Specific sheet either by position or by name
#' str(xlsx_cells(examples, 2))
#' str(xlsx_cells(examples, "Sheet1"))
#'
#' # The formats of particular cells can be retrieved like this:
#'
#' Sheet1 <- xlsx_cells(examples)$Sheet1
#' formats <- xlsx_formats(examples)
#'
#' formats$local$font$bold[Sheet1$local_format_id]
#' formats$style$font$bold[Sheet1$style_format]
#'
#' # To filter for cells of a particular format, first filter the formats to get
#' # the relevant indices, and then filter the cells by those indices.
#' bold_indices <- which(formats$local$font$bold)
#' Sheet1[Sheet1$local_format_id %in% bold_indices, ]
#'
#' # In-cell formatting is available in the `character_formatted` column as a
#' # data frame, one row per substring.
#' xlsx_cells(examples)$character_formatted[77]
xlsx_cells <- function(path, sheets = NA, check_filetype = TRUE) {
  path <- check_file(path)
  sheets <- check_sheets(sheets, path)
  xlsx_cells_(path,
              sheets$sheet_path,
              sheets$name,
              sheets$comments_path)
}
