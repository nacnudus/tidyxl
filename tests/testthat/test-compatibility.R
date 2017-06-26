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

test_that("warns about default styles when no cellStyleXfs defined", {
  expect_warning(tidy_xlsx("haskell.xlsx"),
                 "Default styles used \\(cellStyleXfs is not defined\\)")
})

test_that("libreoffice 'true' and 'false' are interpreted as bool", {
  expect_equal(tidy_xlsx("libreoffice_bool.xlsx")$formats$local$font$bold,
               c(FALSE, TRUE))
})

test_that("libreoffice 'applyFill' defaults to true (it is never defined)", {
  x <- tidy_xlsx("libreoffice-fill.xlsx")
  expect_equal(x$formats$local$fill$patternFill$patternType[13], "solid")
})
