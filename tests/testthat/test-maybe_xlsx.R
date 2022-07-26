context("maybe_xlsx")

test_that("non-xlsx files are detected", {
  expect_equal(maybe_xlsx("examples.xlsx"), TRUE)
  expect_equal(maybe_xlsx("examples.xlsm"), TRUE)
  expect_equal(maybe_xlsx("examples.xltx"), TRUE)
  expect_equal(maybe_xlsx("examples.xltm"), TRUE)
  expect_equal(maybe_xlsx("examples.xlsb"), TRUE) # Unfortunately xlsb look like xlsx
  expect_equal(maybe_xlsx("examples.xls"), FALSE)
  expect_error(xlsx_cells("examples.xlsx", check_filetype = FALSE), NA)
  expect_error(xlsx_cells("examples.xlsm", check_filetype = FALSE), NA)
  expect_error(xlsx_cells("examples.xltx", check_filetype = FALSE), NA)
  expect_error(xlsx_cells("examples.xltm", check_filetype = FALSE), NA)
  expect_error(xlsx_cells("examples.xlsb", check_filetype = FALSE), "Couldn't find.*")
  expect_error(xlsx_cells("examples.xls"), "The file format*")
})
