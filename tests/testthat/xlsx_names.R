context("xlsx_names()")

library(tibble)

test_that("xlsx_names() doesn't fail when there aren't any names", {
  expect_error(xlsx_names("./default.xlsx"), NA)
})

test_that("xlsx_names() refers to the correct sheet names", {
  expect_equal(tidyxl::xlsx_names("./examples.xlsx")$sheet_name[2], "E09904.2")
})

test_that("xlsx_names() determines correctly whether a name is a range", {
  expect_equal(tidyxl::xlsx_names("./examples.xlsx")$is_range[1:5],
               c(FALSE, TRUE, TRUE, FALSE, TRUE))
})
