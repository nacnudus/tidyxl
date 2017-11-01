context("formats")

examples <- "examples.xlsx"
cells <- xlsx_cells(examples, "Sheet1")
style <- function(address) {cells[cells$address == address, ]$style_format}
local_id <- function(address) {cells[cells$address == address, ]$local_format_id}
styles <- xlsx_formats(examples)$style
locals <- xlsx_formats(examples)$local

test_that("number formats inherit from styles", {
  expect_equal(unname(styles$numFmt[style("A12")]), "0%")
  expect_equal(locals$numFmt[local_id("A12")], "0%")
})

test_that("number styles are independent of local formats", {
  expect_equal(unname(styles$numFmt[style("A6")]), "General")
  expect_equal(locals$numFmt[local_id("A6")], "d-mmm-yy")
})

test_that("the 'General' number format is parsed correctly", {
  expect_equal(unname(styles$numFmt[style("A1")]), "General")
  expect_equal(locals$numFmt[local_id("A1")], "General")
})

test_that("custom formats are parsed correctly", {
  expect_equal(locals$numFmt[local_id("A7")], "[$-1409]d\\ mmmm\\ yyyy;@")
})

test_that("font formats inherit from styles", {
  expect_equal(unname(styles$font$name[style("A26")]), "Arial")
  expect_equal(locals$font$name[local_id("A26")], "Arial")
})

test_that("font styles are independent of local formats", {
  expect_equal(unname(styles$font$name[style("A27")]), "Arial")
  expect_equal(locals$font$name[local_id("A27")], "Calibri")
})

test_that("font names are parsed correctly", {
  expect_equal(unname(styles$font$name[style("A27")]), "Arial")
  expect_equal(locals$font$name[local_id("A27")], "Calibri")
})

test_that("font italics are parsed correctly", {
  expect_equal(locals$font$italic[local_id("A26")], FALSE)
  expect_equal(locals$font$italic[local_id("A27")], TRUE)
})

test_that("font bold is parsed correctly", {
  expect_equal(locals$font$bold[local_id("A27")], FALSE)
  expect_equal(locals$font$bold[local_id("A28")], TRUE)
})

test_that("font underline is parsed correctly", {
  expect_equal(locals$font$underline[local_id("A28")], NA_character_)
  expect_equal(locals$font$underline[local_id("A29")], "single")
  expect_equal(locals$font$underline[local_id("A30")], "double")
  expect_equal(locals$font$underline[local_id("A133")], "singleAccounting")
  expect_equal(locals$font$underline[local_id("A134")], "doubleAccounting")
})

test_that("font subscript and superscript is parsed correctly", {
  expect_equal(locals$font$vertAlign[local_id("A135")], NA_character_)
  expect_equal(locals$font$vertAlign[local_id("A136")], "superscript")
  expect_equal(locals$font$vertAlign[local_id("A137")], "subscript")
})

test_that("font sizes are parsed correctly", {
  expect_equal(locals$font$size[local_id("A39")], 11)
  expect_equal(locals$font$size[local_id("A40")], 12)
})

test_that("fonts colours are parsed correctly", {
  expect_equal(locals$font$color$rgb[local_id("A1")], NA_character_)
  expect_equal(locals$font$color$theme[local_id("A1")], NA_integer_)
  expect_equal(locals$font$color$indexed[local_id("A1")], NA_integer_)
  expect_equal(locals$font$color$tint[local_id("A1")], NA_real_)
  expect_equal(locals$font$color$rgb[local_id("A35")], "FFFF0000")
  expect_equal(locals$font$color$theme[local_id("A35")], NA_integer_)
  expect_equal(locals$font$color$indexed[local_id("A35")], NA_integer_)
  expect_equal(locals$font$color$tint[local_id("A35")], NA_real_)
  expect_equal(locals$font$color$rgb[local_id("A36")], "FF4F81BD")
  expect_equal(locals$font$color$theme[local_id("A36")], 5L)
  expect_equal(locals$font$color$indexed[local_id("A36")], NA_integer_)
  expect_equal(locals$font$color$tint[local_id("A36")], NA_real_)
  expect_equal(locals$font$color$rgb[local_id("A138")], "FFF79646")
  expect_equal(locals$font$color$theme[local_id("A138")], 10L)
  expect_equal(locals$font$color$indexed[local_id("A138")], NA_integer_)
  expect_equal(locals$font$color$tint[local_id("A138")], -0.2499771) # Excel's precision
})
