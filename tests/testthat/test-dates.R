context("Dates")

test_that("1900 and 1904 dates are parsed correctly", {
  expect_equal(tidy_xlsx("./1900.xlsx")$data$Sheet1$date[1], as.POSIXct("1998-07-05", tz = "UTC"))
  expect_equal(tidy_xlsx("./1904.xlsx")$data$Sheet1$date[1], as.POSIXct("1998-07-05", tz = "UTC"))
})
