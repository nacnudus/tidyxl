# Tests adapted from tidyverse/readxl

context("Inline strings")

test_that("can read sheets with inlineStr", {
  # Original source: http://our.componentone.com/wp-content/uploads/2011/12/TestExcel.xlsx
  # These appear to come from LibreOffice 4.2.7.2.
  x <- xlsx_cells("inlineStr.xlsx")$character
  expect_equal(x[14], "RQ11610")
})

test_that("can read sheets with formatted inline strings", {
  # Original source: http://our.componentone.com/wp-content/uploads/2011/12/TestExcel.xlsx
  # These appear to come from LibreOffice 4.2.7.2.
  cells <- xlsx_cells("inline-formatted-string.xlsx")
  index <- 64
  expect_equal(cells$is_blank[index], FALSE)
  expect_equal(cells$data_type[index], "character")
  expect_failure(expect_equal(cells$character_formatted[index], NULL))
})

test_that("does not crash on phonetic strings", {
  # https://github.com/nacnudus/tidyxl/issues/30
  expect_error(xlsx_cells("phonetic.xlsx"), NA)
})
