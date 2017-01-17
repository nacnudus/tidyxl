# Tests adapted from hadley/readxl

context("Compatability")

expect_that("can read document from google doc", {
  # All I want to check is that it doesn't crash, but there's no expect_survival
  expect_silent(tidy_xlsx("iris-google-doc.xlsx"))
})
