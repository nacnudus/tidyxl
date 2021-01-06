# Tests adapted from tidyverse/readxl

context("Rich text")

test_that("rich text strings are handled in stringtable", {
  rt <- xlsx_cells("richtext-coloured.xlsx")$character

  for (i in 1:4)
    expect_equal(rt[i], "abcd")

  expect_equal(rt[5], "tvalrval1rval2")
  for (i in 6:8)
    expect_equal(rt[i], "rval1rval2")
})

test_that("rich text inside inlineStr", {
  # Original source: http://our.componentone.com/wp-content/uploads/2011/12/TestExcel.xlsx
  # Modified to have Excel-safe mixed use of <r> and <t>
  x <- xlsx_cells("inlineStr2.xlsx", "Requirements")$character
  expect_equal(x[14], "RQ11610")
  expect_equal(x[1], "NNNN")
  expect_equal(x[2], "BeforeHierarchy")
})

test_that("strings containing escaped hexcodes are read", {
  x <- xlsx_cells("new_line_errors.xlsx", "Blad1")$character
  expect_false(grepl("_x000D_", x[2]))
  expect_equal(substring(x[2], 20, 21), "\u000d\r")
  expect_equal(substring(x[3],11,19), "\"_x000D_\"")
})

test_that("Unicode text is read", {
  shruggie <- "¯\\_(ツ)_/¯"
  Encoding(shruggie) <- "UTF-8"
  expect_equal(xlsx_cells("examples.xlsx")$character[186], shruggie)
})

test_that("Text is encoded as UTF-8", {
  x <- xlsx_cells("utf8.xlsx")
  expect_equal(Encoding(x$comment), "UTF-8")
  expect_equal(Encoding(x$character), "UTF-8")
})

test_that("Unicode in values of formulas is UTF-8", {
  x <- xlsx_cells("german_umlaute_utf8.xlsx")
  expect_equal(Encoding(x$character[-1]), rep("UTF-8", 9))
})

test_that("Unicode in comments is UTF-8", {
  x <- xlsx_cells("german_umlaute_utf8.xlsx")
  expect_equal(Encoding(x$comment[-1]), rep("UTF-8", 9))
})
