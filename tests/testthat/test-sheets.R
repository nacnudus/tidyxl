# Tests adapted from hadley/readxl

context("Sheets")

test_that("xlsx_sheets returns utf-8 encoded text", {
  sheets <- tidyxl:::xlsx_sheets("utf8-sheets.xlsx")$name
  expect_equal(Encoding(sheets), rep("UTF-8", 2))
  expect_equal(sheets, c("\u00b5", "\u2202"))
})
