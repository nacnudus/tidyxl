context("Dates")

# # openxlsx1900.xlsx is created as follows
# # (https://github.com/tidyverse/readxl/issues/264):
# library(openxlsx)
# dat <- data.frame(A=c(1, 1.5, 1.8))
# wb <- createWorkbook()
# addWorksheet(wb=wb, sheetName = "Sheet1")
# sty <- createStyle(numFmt = "h:mm AM/PM")
# addStyle(wb, sheet=1, style=sty, rows=2:4, cols=1)
# writeData(wb, sheet=1, dat)
# saveWorkbook(wb, "openxlsx1900.xlsx", overwrite = TRUE)

test_that("1900 and 1904 dates are parsed correctly", {
  expect_equal(xlsx_cells("./1900.xlsx")$date[1], as.POSIXct("1998-07-05", tz = "UTC"))
  expect_equal(xlsx_cells("./1900.xlsx")$date[3], as.POSIXct("1900-02-28", tz = "UTC"))
  expect_equal(xlsx_cells("./1904.xlsx")$date[1], as.POSIXct("1998-07-05", tz = "UTC"))
  # This is the only way to construct an NA POSIXct with timezone
  NA_POSIXct <- as.POSIXct(NA)
  attr(NA_POSIXct, "tzone") <- "UTC"
  expect_warning(expect_equal(xlsx_cells("./1900-02-29.xlsx")$date[1], NA_POSIXct),
                 "NA inserted for impossible 1900-02-29 datetime: 'Sheet1'!A1")
  expect_equal(xlsx_cells("./openxlsx1900.xlsx")$date[2],
               as.POSIXct("1900-01-01", tz = "UTC"))
})

test_that("Colours in custom number formats aren't detected as date formats", {
  x <- xlsx_cells("./rainbow.xlsx")
  expect_equal(x[x$col == 2, ]$data_type, rep("numeric", 8))
})
