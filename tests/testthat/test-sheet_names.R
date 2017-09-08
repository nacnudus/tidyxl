context("xlsx_sheet_names()")

test_that("sheet names are returned correctly", {
  expect_equal(xlsx_sheet_names("./examples.xlsx")[1:2],
               c("Sheet1", "1~`!@#$%^&()_-+={}|;\"'<,>.£¹÷¦°"))
})
