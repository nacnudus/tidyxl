context("is_range()")

test_that("is_range() correctly identifies ranges", {
  expect_equal(is_range("A1"),           TRUE)
  expect_equal(is_range("Sheet1!A1"),    TRUE)
  expect_equal(is_range("[0]Sheet1!A1"), TRUE)
  expect_equal(is_range("A1,A2"),        TRUE)
  expect_equal(is_range("A:A 3:3"),      TRUE)
  expect_equal(is_range("MAX(A1,2)"),    FALSE)
})
