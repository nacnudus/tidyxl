#' Import xlsx (Excel) spreadsheet data and formatting into tidy structures.
#'
#' Tidyxl imports data from spreadsheets without coercing it into a rectangle,
#' and retains information encoded in cell formatting (e.g. font/fill/border).
#' This data structure is compatible with the 'unpivotr' package for
#' recognising information expressed by relative cell positions and cell
#' formatting, and re-expressing it in a tidy way.
#'
#' Two functions are provided: \code{\link{contents}} for importing cell contents,
#' and \code{\link{formats}} for importing cell formatting.  To see how to link
#' contents and formatting, see \code{\link{formats}}.
#'
#' The pipe \code{\%>\%} function from the magrittr package is re-exported, and
#' the tidyr package is attached so that the output of \code{\link{contents}} is
#' printed nicely (as a tibble rather than a data frame).
"_PACKAGE"
#> [1] "_PACKAGE"
