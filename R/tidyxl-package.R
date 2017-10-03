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
#' @section Spreadsheet functions:
#' * [tidyxl::tidy_xlsx()] Import cells and formatting from an xlsx file.
#' * [tidyxl::xlsx_sheet_names()] List the names of sheets in an xlsx file.
#' * [tidyxl::xlsx_names()] Import names and definitions of named ranges (aka named
#' * [tidyxl::is_range()] Test whether a 'name' from [tidyxl::xlsx_names()]
#' refers to a range or not.
#' formulas, defined names) from an xlsx file.
#' * [tidyxl::xlsx_validation()] Import cell input validation rules (e.g.
#' 'must be from this drop-down list') from an xlsx file.
#'
#' @section Formula functions:
#' * [tidyxl::xlex()] Tokenise (lex) an Excel formula.
#' * [tidyxl::plot_xlex()] Draw a simple tree plot of a parse tree from
#' [tidyxl::xlex()]
#' * [tidyxl::xlex_edges()] Utility function used by [tidyxl::plot_xlex()]
#' * [tidyxl::xlex_vertices()] Utility function used by [tidyxl::plot_xlex()]
#' * [tidyxl::xlex_igraph()] Utility function used by [tidyxl::plot_xlex()]
#'
#' @docType package
#' @name tidyxl
NULL
