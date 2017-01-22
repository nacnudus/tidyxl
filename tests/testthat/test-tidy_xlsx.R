context("tidy_xlsx()")

test_that("warns of missing sheets", {
  expect_warning(expect_error(tidy_xlsx("./examples.xlsx", c(NA, NA)),"All elements of argument 'sheets' were discarded."),  "Argument 'sheets' included NAs, which were discarded.")
  expect_error(tidy_xlsx("./examples.xlsx", "foo"), "Sheet\\(s\\) not found: \"foo\"")
  expect_error(tidy_xlsx("./examples.xlsx", c("foo", "bar")), "Sheet\\(s\\) not found: \"foo\", \"bar\"")
})

