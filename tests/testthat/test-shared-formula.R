context("Shared formulas")

test_that("Shared formulas are propogated correctly", {
  formulas <- tidyxl::tidy_xlsx("./examples.xlsx")$data[[1]]$formula
  expect_equal(formulas[41], "$A$18+1")
  expect_equal(formulas[42], "A20+2")
  expect_equal(formulas[195], "C3&\"C1\"\"\"")
  expect_equal(formulas[199], "LOG12")
  expect_equal(formulas[203], "LOG10(1)")
  expect_equal(formulas[224], "A1A")
  expect_equal(formulas[228], "E09904.2!A3")
  formulas <- tidyxl::tidy_xlsx("./formulas.xlsx")$data[[1]]$formula
  expect_equal(formulas[3], "N($A2)")
})

