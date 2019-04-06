#' @title Import xlsx (Excel) formatting definitions.
#'
#' @description
#' `xlsx_formats()` imports formatting definitions from spreadsheets.  The
#' structure is a nested list, e.g. `bold` is a vector within the list `font`,
#' which is within the list `local`, which is within the list
#' returned by `xlsx_formats()`.  You can look up a cell's formatting by
#' indexing the bottom-level vectors.  See 'Details' for examples.
#'
#' @param path Path to the xlsx file.
#' @param check_filetype Logical. Whether to check that the filetype is xlsx (or
#' xlsm) by looking at the file itself, rather than using the filename
#' extension.
#'
#' @return
#' A nested list of vectors, beginning at the top level with `$style` and
#' `$local`, then drilling down to the vectors that hold the definitions.  E.g.
#' `my_formats$local$font$size`.
#'
#' @details
#' There are two types of formatting: 'style' formatting, such as Excel's
#' built-in styles 'normal', 'bad', etc., and 'local' formatting, which
#' overrides the style.  These are returned in the `$style` and `$local`
#' sublists of `xlsx_formats()`, with identical structures.
#'
#' To look up the local formatting of a given cell, take the cell's
#' `local_format_id` value (`my_cells$Sheet1[1, "local_format_id"]`), and use it
#' as an index into the format structure.  E.g.  to look up the font size,
#' `my_formats$local$font$size[local_format_id]`.  To see all available formats,
#' type `str(my_formats$local)`.
#'
#' Colours may be recorded in any of three ways: a hexadecimal RGB string with
#' or without alpha, an 'indexed' colour, and an index into a 'theme'.
#' `xlsx_formats()` dereferences 'indexed' and 'theme' colours to their
#' hexadecimal RGB string representation, and standardises all RGB strings to
#' have an alpha channel in the first two characters.  The 'index' and the
#' 'theme' name are still provided.  To filter by an RGB string, you could  look
#' up the RGB values in a spreadsheet program (e.g. Excel, LibreOffice,
#' Gnumeric), and use the [grDevices::rgb()] function to convert these to a
#' hexadecimal string.
#'
#'   ```
#'   A <- 1; R <- 0.5; G <- 0; B <- 0
#'   rgb(A, R, G, B)
#'   # [1] "#FF800000"
#'   ```
#'
#' @export
#' @examples
#' examples <- system.file("extdata/examples.xlsx", package = "tidyxl")
#' str(xlsx_formats(examples))
#'
#' # The formats of particular cells can be retrieved like this:
#'
#' cells <- xlsx_cells(examples)
#' formats <- xlsx_formats(examples)
#'
#' formats$local$font$bold[cells$local_format_id]
#' formats$style$font$bold[cells$style_format]
#'
#' # To filter for cells of a particular format, first filter the formats to get
#' # the relevant indices, and then filter the cells by those indices.
#' bold_indices <- which(formats$local$font$bold)
#' cells[cells$local_format_id %in% bold_indices, ]
xlsx_formats <- function(path, check_filetype = TRUE) {
  path <- check_file(path)
  xlsx_formats_(path)
}
