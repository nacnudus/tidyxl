context("tidy_xlsx()")

test_that("sheet order is correct", {
  expect_equal(names(tidy_xlsx("./sheet-order.xlsx")$data),
               c("Sheet1", "Sheet2", "Sheet3"))
  expect_equal(tidy_xlsx("./sheet-order.xlsx", 2)$data[[1]]$numeric, 3)
})
