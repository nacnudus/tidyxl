context("xlsx_sheet_names()")

test_that("sheet names are returned correctly", {
  expect_equal(xlsx_sheet_names("./examples.xlsx")[1:2],
               c("Sheet1", "1~`!@#$%^&()_-+={}|;\"'<,>.£¹÷¦°"))
})

test_that("sheet name order is correct", {
  expect_equal(unique(xlsx_sheet_names("./lu-performance-data-almanac.xlsx")[1:3]),
               c("All Sheets", "Cover Page", "Contents"))
})
