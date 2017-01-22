# Tests adapted from hadley/readxl

context("Compatability")

test_that("can import examples.xlsx", {
  expect_error(x <- tidy_xlsx("./examples.xlsx"), NA)
})

test_that("can import examples-gnumeric.xlsx", {
  expect_error(x <- tidy_xlsx("./examples-gnumeric.xlsx"), NA)
})

test_that("can read document from google doc", {
  # All I want to check is that it doesn't crash, but there's no expect_survival
  expect_silent(tidy_xlsx("iris-google-doc.xlsx"))
})

