#' @useDynLib tidyxl
#' @importFrom Rcpp sourceCpp
NULL

#' tidyxl: Import xlsx (Excel) spreadsheet data and formatting into tidy structures.
#'
#' Tidyxl imports data from spreadsheets without coercing it into a rectangle,
#' and retains information encoded in cell formatting (e.g. font/fill/border).
#' This data structure is compatible with the 'unpivotr' package for
#' recognising information expressed by relative cell positions and cell
#' formatting, and re-expressing it in a tidy way.
#'
#' * [tidyxl::xlsx_cells()] Import cells from an xlsx file.
#' * [tidyxl::xlsx_formats()] Import formatting from an xlsx file.
#' * [tidyxl::xlsx_sheet_names()] List the names of sheets in an xlsx file.
#' * [tidyxl::xlsx_names()] Import names and definitions of named ranges (aka
#'   'named formulas', 'defined names') from an xlsx file.
#' * [tidyxl::is_range()] Test whether a 'name' from [tidyxl::xlsx_names()]
#'   refers to a range or not.
#' * [tidyxl::xlsx_validation()] Import cell input validation rules (e.g.
#'   'must be from this drop-down list') from an xlsx file.
#' * [tidyxl::xlsx_colour_standard()] A data frame of standard colour names and
#'   their RGB values.
#' * [tidyxl::xlsx_colour_theme()] Imports a data frame of theme colour names
#'   and their RGB values from an xlsx file.
#' * [tidyxl::xlex()] Tokenise (lex) an Excel formula.
#'
#' @docType package
#' @name tidyxl
NULL
