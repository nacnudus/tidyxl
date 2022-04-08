context("xlsx_cells")

test_that("warns of missing sheets", {
  expect_warning(expect_error(xlsx_cells("./examples.xlsx", c(NA, NA)),"All elements of argument 'sheets' were discarded."),  "Argument 'sheets' included NAs, which were discarded.")
  expect_error(xlsx_cells("./examples.xlsx", "foo"), "Sheets not found: \"foo\"")
  expect_error(xlsx_cells("./examples.xlsx", c("foo", "bar")), "Sheets not found: \"foo\", \"bar\"")
  expect_error(xlsx_cells("./examples.xlsx", 6), "Only 5 sheet\\(s\\) found.")
  expect_warning(xlsx_cells("./examples.xlsx", 3), "Only worksheets \\(not chartsheets\\) were imported")
  expect_error(xlsx_cells("./examples.xlsx", TRUE), "Argument `sheet` must be either an integer or a string.")
})

test_that("finds named sheets", {
  expect_error(xlsx_cells("./examples.xlsx", "Sheet1"), NA)
})

test_that("sheet order is correct", {
  expect_equal(unique(xlsx_cells("./sheet-order.xlsx")$sheet),
               c("Sheet1", "Sheet2", "Sheet3"))
  expect_equal(xlsx_cells("./sheet-order.xlsx", 2)$numeric, 3)
})

test_that("gracefully fails on missing files", {
  expect_error(xlsx_cells("foo.xlsx"), "'foo\\.xlsx' does not exist in current working directory \\('.*'\\).")
})

test_that("allows user interruptions", {
  # This is just to test the branch that waits for interruptions.  No
  # interruption is attempted
  expect_error(xlsx_cells("./thousand.xlsx"), NA)
})

test_that("The highest cell address is parsed", {
  expect_error(xlsx_cells("./xfd1048576.xlsx"), NA)
})

test_that("array formulas are detected as such", {
  cells <- xlsx_cells("./examples.xlsx")
  expect_equal(cells$is_array[41], FALSE)
  expect_equal(cells$is_array[43], TRUE)
  expect_equal(cells$is_array[45], TRUE)
})

test_that("include_blank_cells works", {
  cells <- xlsx_cells("./examples.xlsx", include_blank_cells = FALSE)
  blanks <- cells[cells$is_blank, ]
  non_blanks <- cells[!cells$is_blank, ]
  unpopulated_blanks <- cells[is.na(cells$row), ]
  expect_equal(nrow(blanks), 0L)
  expect_gt(nrow(non_blanks), 0L)
  expect_equal(nrow(unpopulated_blanks), 0L)
})

test_that("raw cell values are returned", {
  cells <- xlsx_cells("./examples.xlsx", sheet = "Sheet1")
  cells <- cells[cells$col == 1 & cells$row %in% c(1, 2, 4, 6, 9), ]
  expect_equal(cells$content, c("#DIV/0!", "1", "1337", "42046", "106"))
})
