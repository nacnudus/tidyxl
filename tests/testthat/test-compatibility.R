# Tests adapted from hadley/readxl

context("Compatability")

test_that("can import examples.xlsx", {
  expect_error(x <- tidy_xlsx("./examples.xlsx"), NA)
})

test_that("can import examples-gnumeric.xlsx", {
  expect_error(x <- tidy_xlsx("./examples-gnumeric.xlsx"), NA)
})

test_that("can read document from google doc", {
  expect_error(tidy_xlsx("iris-google-doc.xlsx"), NA)
})

test_that("gives informative error for a JMP export", {
  expect_error(tidy_xlsx("jmp.xlsx"), "Invalid cell: lacks 'r' attribute")
})
