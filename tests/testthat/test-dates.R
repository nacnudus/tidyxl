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
  expect_equal(tidy_xlsx("./1900.xlsx")$data$Sheet1$date[1], as.POSIXct("1998-07-05", tz = "UTC"))
  expect_equal(tidy_xlsx("./1900.xlsx")$data$Sheet1$date[3], as.POSIXct("1900-02-28", tz = "UTC"))
  expect_equal(tidy_xlsx("./1904.xlsx")$data$Sheet1$date[1], as.POSIXct("1998-07-05", tz = "UTC"))
  # This is the only way to construct an NA POSIXct with timezone
  NA_POSIXct <- as.POSIXct(NA)
  attr(NA_POSIXct, "tzone") <- "UTC"
  expect_warning(expect_equal(tidy_xlsx("./1900-02-29.xlsx")$data$Sheet1$date[1], NA_POSIXct),
                 "NA inserted for impossible 1900-02-29 datetime: 'Sheet1'!A1")
  expect_equal(tidy_xlsx("./openxlsx1900.xlsx")$data$Sheet1$date[2],
               as.POSIXct("1900-01-01", tz = "UTC"))
})

test_that("Custom number formats using [Red] aren't detected as dates", {
  expect_equal(tidy_xlsx("./examples.xlsx")$data$Sheet1$data_type[188], "numeric")
})
