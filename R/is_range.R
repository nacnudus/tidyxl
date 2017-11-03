#' Test that Excel formulas are ranges
#'
#' @description `is_range()` tests whether or not an Excel formula is a range.
#' A formula like `A1` is a range, whereas a formula like `MAX(A1,2)` is not.
#' Formulas are not evaluated, so it returns `FALSE` for formulas that would
#' eventually resolve to arrange (e.g.  `INDEX(A1:A10,2)`) but that are not
#' immediately a range.
#'
#' @param x character vector of formulas
#' @export
#' @examples
#' x <- c("A1", "Sheet1!A1", "[0]Sheet1!A1", "A1,A2", "A:A 3:3", "MAX(A1,2)")
#' is_range(x)
is_range <- function(x) {
  is_range_scalar <- function(y) {
    # Contiguous sequence of ref[(union or intersection operator)ref]*
    tokens <- xlex(y)
    tokens <- tokens[tokens$type != "sheet", ]
    # Now all refs should be comma-or-space-separated
    len <- nrow(tokens)
    odds <- seq.int(1, len, by = 2)
    if (len %% 2 == 0 && len >= 2) {
      evens <- seq.int(2, len, by = 2) # nocov start
      out <- all(tokens$type[odds] == "ref") && # covr doesn't realise these ARE tested
               all(tokens$token[evens] %in% c(",", " ")) # nocov end
    } else {
      out <- all(tokens$type[odds] == "ref")
    }
    out
  }
  vapply(as.list(x), is_range_scalar, logical(1))
}
