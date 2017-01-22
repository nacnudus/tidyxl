context("Dates")

test_that("1900 and 1904 dates are parsed correctly", {
  expect_equal(tidy_xlsx("./1900.xlsx")$data$Sheet1$date[1], as.POSIXct("1998-07-05", tz = "UTC"))
  expect_equal(tidy_xlsx("./1904.xlsx")$data$Sheet1$date[1], as.POSIXct("1998-07-05", tz = "UTC"))
})

test_that("A warning is given for 1900-system dates before 1 March 1900", {
  expect_warning(tidy_xlsx("./date-bug.xlsx"), "Dates before 1 March 1900 are off by one, due to Excel's famous bug.")
})
