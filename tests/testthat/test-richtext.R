# Tests adapted from hadley/readxl

context("Richtext")

test_that("rich text strings are handled in stringtable", {
  rt <- contents("richtext-coloured.xlsx")$Sheet1$character

  for (i in 1:4)
    expect_equal(rt[i], "abcd")

  expect_equal(rt[5], "tvalrval1rval2")
  for (i in 6:8)
    expect_equal(rt[i], "rval1rval2")
})

test_that("rich text inside inlineStr", {
  # Original source: http://our.componentone.com/wp-content/uploads/2011/12/TestExcel.xlsx
  # Modified to have Excel-safe mixed use of <r> and <t>
  x <- contents("inlineStr2.xlsx")$Requirements$character
  expect_equal(x[14], "RQ11610")
  expect_equal(x[1], "NNNN")
  expect_equal(x[2], "BeforeHierarchy")
})

test_that("strings containing escaped hexcodes are read", {
  x <- contents("new_line_errors.xlsx")$Blad1$character
  expect_false(grepl("_x000D_", x[2]))
  expect_equal(substring(x[2], 20, 21), "\u000d\r")
  expect_equal(substring(x[3],11,19), "\"_x000D_\"")
})
