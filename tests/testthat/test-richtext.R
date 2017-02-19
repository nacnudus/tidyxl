# Tests adapted from hadley/readxl

context("Rich text")

test_that("rich text strings are handled in stringtable", {
  rt <- tidy_xlsx("richtext-coloured.xlsx")$data$Sheet1$character

  for (i in 1:4)
    expect_equal(rt[i], "abcd")

  expect_equal(rt[5], "tvalrval1rval2")
  for (i in 6:8)
    expect_equal(rt[i], "rval1rval2")
})

test_that("rich text inside inlineStr", {
  # Original source: http://our.componentone.com/wp-content/uploads/2011/12/TestExcel.xlsx
  # Modified to have Excel-safe mixed use of <r> and <t>
  x <- tidy_xlsx("inlineStr2.xlsx")$data$Requirements$character
  expect_equal(x[14], "RQ11610")
  expect_equal(x[1], "NNNN")
  expect_equal(x[2], "BeforeHierarchy")
})

test_that("strings containing escaped hexcodes are read", {
  x <- tidy_xlsx("new_line_errors.xlsx")$data$Blad1$character
  expect_false(grepl("_x000D_", x[2]))
  expect_equal(substring(x[2], 20, 21), "\u000d\r")
  expect_equal(substring(x[3],11,19), "\"_x000D_\"")
})

test_that("Unicode text is read", {
  shruggie <- "¯\\_(ツ)_/¯"
  Encoding(shruggie) <- "UTF-8"
  expect_equal(tidy_xlsx("examples.xlsx")$data$Sheet1$character[186], shruggie)
})

test_that("Text is encoded as UTF-8", {
  x <- tidy_xlsx("utf8.xlsx")
  expect_equal(Encoding(x$data$Sheet1$comment), "UTF-8")
  expect_equal(Encoding(x$data$Sheet1$character), "UTF-8")
})
