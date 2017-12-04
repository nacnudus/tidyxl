#' Determine file format
#'
#' Whether a file may be be xlsx or xlsm (rather than xls), based on the
#' [file signature](https://en.wikipedia.org/wiki/List_of_file_signatures) or
#' "magic number", rather than the filename extension.
#'
#' Only 'maybe', not 'is', because the xlsx magic number is common to all zip
#' files, not specific to xlsx files.  The inverse, 'is_xls' isn't possible
#' either, because the xls magic number is common to other Microsoft Office
#' files such as .doc and .ppt.
#'
#' This uses some logic from Jenny Bryan's commit to the
#' [readxl](https://github.com/tidyverse/readxl/commit/ff071a4758da3677568362daff52e419c4e0cdfe)
#' package.
#'
#' @param path File path
#'
#' @return Logicial
#'
#' @export
#' @examples
#' examples_xlsx <- system.file("extdata/examples.xlsx", package = "tidyxl")
#' examples_xlsm <- system.file("extdata/examples.xlsm", package = "tidyxl")
#' examples_xlsb <- system.file("extdata/examples.xlsb", package = "tidyxl")
#' examples_xls <- system.file("extdata/examples.xls", package = "tidyxl")
#'
#' maybe_xlsx(examples_xlsx)
#' maybe_xlsx(examples_xlsm)
#' maybe_xlsx(examples_xlsb)
#' maybe_xlsx(examples_xls)
maybe_xlsx <- function(path) {
  if (!file.exists(path)) {
    stop("'", path, "' does not exist",
      if (!is_absolute_path(path))
        paste0(" in current working directory ('", getwd(), "')"),
      ".",
      call. = FALSE)
  }
  signature <- readBin(path, n = 8, what = "raw")
  xlsx_sig <- as.raw(c("0x50", "0x4B", "0x03", "0x04"))
  identical(signature[1:4], xlsx_sig)
}
