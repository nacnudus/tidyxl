# Tests adapted from hadley/readxl

context("Compatability")

test_that("can import examples.xlsx", {
  expect_error(xlsx_cells("./examples.xlsx"), NA)
  expect_error(xlsx_formats("./examples.xlsx"), NA)
})

test_that("can import examples-gnumeric.xlsx", {
  expect_error(xlsx_cells("./examples-gnumeric.xlsx"), NA)
  expect_error(xlsx_formats("./examples-gnumeric.xlsx"), NA)
})

test_that("can read document from google doc", {
  expect_error(xlsx_cells("iris-google-doc.xlsx"), NA)
  expect_error(xlsx_formats("iris-google-doc.xlsx"), NA)
})

test_that("gives informative error for a JMP export", {
  expect_error(xlsx_cells("jmp.xlsx"), "Invalid row or cell: lacks 'r' attribute")
})

test_that("warns about default styles when no cellStyleXfs defined", {
  expect_warning(xlsx_cells("haskell.xlsx"),
                 "Default styles used \\(cellStyleXfs is not defined\\)")
})

test_that("libreoffice 'true' and 'false' are interpreted as bool", {
  expect_equal(xlsx_formats("libreoffice-bool.xlsx")$local$font$bold,
               c(FALSE, TRUE))
})

test_that("libreoffice 'applyFill' defaults to true (it is never defined)", {
  x <- xlsx_formats("libreoffice-fill.xlsx")
  expect_equal(x$local$fill$patternFill$patternType[13], "solid")
})

test_that("Missing styles don't cause crashes", {
  # The problem was that, if not all styles in a range e.g. 1:36, were present,
  # the name lookup went beyond the bounds of an array.  It should look up a
  # map instead.
  expect_error(xlsx_formats("libreoffice-missing-styles.xlsx"), NA)
})

test_that("Sheet paths like /xl/worksheets/sheet1.xml work", {
  expect_error(xlsx_cells("EvaluacionCensal_Secundaria_SEGUNDO_14112018_160622.xlsx"), NA)
})

test_that("applyNumberFormat etc. are ignored when no format is provided", {
  expect_error(xlsx_cells("EvaluacionCensal_Secundaria_SEGUNDO_14112018_160622.xlsx"), NA)
})
