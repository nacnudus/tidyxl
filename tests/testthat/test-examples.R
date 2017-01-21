context("examples.xlsx")

test_that("can import examples.xlsx", {
  expect_error(x <- tidy_xlsx("./examples.xlsx"), NA)
})

