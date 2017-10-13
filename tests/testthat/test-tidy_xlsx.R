context("tidy_xlsx()")

test_that("sheet order is correct", {
  expect_warning(
    expect_equal(names(tidy_xlsx("./sheet-order.xlsx")$data),
                c("Sheet1", "Sheet2", "Sheet3")),
  "'tidy_xlsx\\(\\)' is deprecated.\nUse 'xlsx_cells\\(\\)' or 'xlsx_formats\\(\\)' instead.")
  expect_warning(
    expect_equal(tidy_xlsx("./sheet-order.xlsx", 2)$data[[1]]$numeric, 3),
  "'tidy_xlsx\\(\\)' is deprecated.\nUse 'xlsx_cells\\(\\)' or 'xlsx_formats\\(\\)' instead.")
})
