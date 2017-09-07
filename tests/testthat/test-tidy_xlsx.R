context("tidy_xlsx()")

test_that("warns of missing sheets", {
  expect_warning(expect_error(tidy_xlsx("./examples.xlsx", c(NA, NA)),"All elements of argument 'sheets' were discarded."),  "Argument 'sheets' included NAs, which were discarded.")
  expect_error(tidy_xlsx("./examples.xlsx", "foo"), "Sheet\\(s\\) not found: \"foo\"")
  expect_error(tidy_xlsx("./examples.xlsx", c("foo", "bar")), "Sheet\\(s\\) not found: \"foo\", \"bar\"")
  expect_error(tidy_xlsx("./examples.xlsx", 5), "Only 3 sheet\\(s\\) found.")
  expect_error(tidy_xlsx("./examples.xlsx", TRUE), "Argument `sheet` must be either an integer or a string.")
})

test_that("finds named sheets", {
  expect_error(tidy_xlsx("./examples.xlsx", "Sheet1"), NA)
})

test_that("sheet order is correct", {
  expect_equal(names(tidy_xlsx("./sheet-order.xlsx")$data),
               c("Sheet1", "Sheet2", "Sheet3"))
  expect_equal(tidy_xlsx("./sheet-order.xlsx", 2)$data[[1]]$numeric, 3)
})

test_that("gracefully fails on missing files", {
  expect_error(tidy_xlsx("foo.xlsx"), "'foo\\.xlsx' does not exist in current working directory \\('.*'\\).")
})

test_that("allows user interruptions", {
  # This is just for code coverage
  expect_error(tidy_xlsx("./thousand.xlsx"), NA)
})

test_that("The highest cell address is parsed", {
  expect_error(tidy_xlsx("./xfd1048576.xlsx"), NA)
})
