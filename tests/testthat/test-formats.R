context("formats")

examples <- "examples.xlsx"

style <- function(address) {
  cells <- xlsx_cells(examples, "Sheet1")
  cells[cells$address == address, ]$style_format
}

local_id <- function(address) {
  cells <- xlsx_cells(examples, "Sheet1")
  cells[cells$address == address, ]$local_format_id
}

styles <- xlsx_formats(examples)$style
locals <- xlsx_formats(examples)$local

test_that("number formats are parsed correctly", {
  expect_equal(unname(styles$numFmt[style("A1")]), "General")
  expect_equal(locals$numFmt[local_id("A1")], "General")
  expect_equal(unname(styles$numFmt[style("A6")]), "General")
  expect_equal(locals$numFmt[local_id("A6")], "d-mmm-yy")
  expect_equal(unname(styles$numFmt[style("A7")]), "General")
  expect_equal(locals$numFmt[local_id("A7")], "[$-1409]d\\ mmmm\\ yyyy;@")
  expect_equal(unname(styles$numFmt[style("A8")]), "General")
  expect_equal(locals$numFmt[local_id("A8")], "yyyy\\ mmmm\\ dddd")
  expect_equal(unname(styles$numFmt[style("A11")]), "General")
  expect_equal(locals$numFmt[local_id("A11")],
               "_-[$£-809]* #,##0.00_-;\\-[$£-809]* #,##0.00_-;_-[$£-809]* \"-\"??_-;_-@_-")
  expect_equal(unname(styles$numFmt[style("A12")]), "0%")
  expect_equal(locals$numFmt[local_id("A12")], "0%")
  expect_equal(unname(styles$numFmt[style("A33")]), "yyyy\\-mm\\-dd\\ hh:mm:ss")
  expect_equal(locals$numFmt[local_id("A33")], "yyyy\\-mm\\-dd\\ hh:mm:ss")
  expect_equal(unname(styles$numFmt[style("A38")]), "yyyy\\-mm\\-dd\\ hh:mm:ss")
  expect_equal(locals$numFmt[local_id("A38")], "yyyy\\-mm\\-dd\\ hh:mm:ss")
  expect_equal(locals$numFmt[local_id("A95")], "\\$#,##0.00;[Red]\\-\\$#,##0.00")
})
