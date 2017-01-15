#' @useDynLib tidyxl
#' @importFrom Rcpp sourceCpp
NULL

#' Import xlsx (Excel) spreadsheet data and formatting into tidy structures.
#'
#' Tidyxl imports data from spreadsheets without coercing it into a rectangle,
#' and retains information encoded in cell formatting (e.g. font/fill/border).
#' This data structure is compatible with the 'unpivotr' package for
#' recognising information expressed by relative cell positions and cell
#' formatting, and re-expressing it in a tidy way.
#'
#' @section Tidyxl functions:
#' Only one function is currently provided: \code{\link{tidy_xlsx}).  Others may
#' be developed for other file formats, e.g. .xls and .ods.
#'
#' @docType package
#' @name tidyxl
