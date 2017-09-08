context("is_range()")

test_that("is_range() correctly identifies ranges", {
  x <- c("A1", "Sheet1!A1", "[0]Sheet1!A1", "A1,A2", "A:A 3:3", "MAX(A1,2)")
  expect_equal(is_range(x), c(TRUE, TRUE, TRUE, TRUE, TRUE, FALSE))
})
