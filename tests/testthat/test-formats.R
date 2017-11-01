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

test_that("fonts colours and colors in general are parsed correctly", {
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

test_that("fill pattern colours are parsed correctly", {
  expect_equal(locals$fill$patternFill$fgColor$rgb[local_id("A69")], NA_character_)
  expect_equal(locals$fill$patternFill$fgColor$rgb[local_id("A70")], "FF00B0F0")
  expect_equal(locals$fill$patternFill$bgColor$rgb[local_id("A69")], NA_character_)
  expect_equal(locals$fill$patternFill$bgColor$rgb[local_id("A70")], "FFFF0000")
})

test_that("fill pattern types are parsed correctly", {
  expect_equal(locals$fill$patternFill$patternType[local_id("A69")], "darkUp")
  expect_equal(locals$fill$patternFill$patternType[local_id("A70")], "solid")
  expect_equal(locals$fill$patternFill$patternType[local_id("A71")], "darkGray")
  expect_equal(locals$fill$patternFill$patternType[local_id("A72")], "mediumGray")
  expect_equal(locals$fill$patternFill$patternType[local_id("A73")], "lightGray")
  expect_equal(locals$fill$patternFill$patternType[local_id("A74")], "gray125")
  expect_equal(locals$fill$patternFill$patternType[local_id("A75")], "gray0625")
  expect_equal(locals$fill$patternFill$patternType[local_id("A76")], "darkHorizontal")
  expect_equal(locals$fill$patternFill$patternType[local_id("A77")], "darkVertical")
  expect_equal(locals$fill$patternFill$patternType[local_id("A78")], "darkDown")
  expect_equal(locals$fill$patternFill$patternType[local_id("A79")], "darkUp")
  expect_equal(locals$fill$patternFill$patternType[local_id("A80")], "darkGrid")
  expect_equal(locals$fill$patternFill$patternType[local_id("A81")], "darkTrellis")
  expect_equal(locals$fill$patternFill$patternType[local_id("A82")], "lightHorizontal")
  expect_equal(locals$fill$patternFill$patternType[local_id("A83")], "lightVertical")
  expect_equal(locals$fill$patternFill$patternType[local_id("A84")], "lightDown")
  expect_equal(locals$fill$patternFill$patternType[local_id("A85")], "lightUp")
  expect_equal(locals$fill$patternFill$patternType[local_id("A86")], "lightGrid")
  expect_equal(locals$fill$patternFill$patternType[local_id("A87")], "lightTrellis")
})

test_that("fill pattern types are consistently NA when not set", {
  expect_equal(locals$fill$patternFill$patternType[local_id("A68")], NA_character_)
  expect_equal(locals$fill$patternFill$patternType[local_id("A88")], NA_character_)
  expect_equal(locals$fill$patternFill$patternType[local_id("A89")], NA_character_)
  expect_equal(locals$fill$patternFill$patternType[local_id("A90")], NA_character_)
  expect_equal(locals$fill$patternFill$patternType[local_id("A91")], NA_character_)
  expect_equal(locals$fill$patternFill$patternType[local_id("A92")], NA_character_)
  expect_equal(locals$fill$patternFill$patternType[local_id("A93")], NA_character_)
  expect_equal(locals$fill$patternFill$patternType[local_id("A94")], NA_character_)
})

test_that("fill gradient colours are parsed correctly", {
  expect_equal(locals$fill$gradientFill$stop1$color$rgb[local_id("A87")], NA_character_)
  expect_equal(locals$fill$gradientFill$stop1$color$rgb[local_id("A88")], "FFFFFFFF")
  expect_equal(locals$fill$gradientFill$stop2$color$rgb[local_id("A88")], "FF4F81BD")
  expect_equal(locals$fill$gradientFill$stop1$color$rgb[local_id("A139")], "FFF79646")
  expect_equal(locals$fill$gradientFill$stop2$color$rgb[local_id("A139")], "FF4F81BD")
})

test_that("fill gradient stop positions are parsed correctly", {
  expect_equal(locals$fill$gradientFill$stop1$position[local_id("A87")], NA_real_)
  expect_equal(locals$fill$gradientFill$stop1$position[local_id("A88")], 0)
  expect_equal(locals$fill$gradientFill$stop2$position[local_id("A88")], 1)
  expect_equal(locals$fill$gradientFill$stop2$position[local_id("A141")], 0.5)
})

test_that("fill gradient types are parsed correctly", {
  expect_equal(locals$fill$gradientFill$type[local_id("A91")], NA_character_)
  expect_equal(locals$fill$gradientFill$type[local_id("A92")], "path")
})

test_that("fill gradient degrees are parsed correctly", {
  expect_equal(locals$fill$gradientFill$degree[local_id("A88")], 90L)
  expect_equal(locals$fill$gradientFill$degree[local_id("A89")], 0L)
  expect_equal(locals$fill$gradientFill$degree[local_id("A92")], NA_integer_)
})

test_that("fill gradient directions are parsed correctly", {
  expect_equal(locals$fill$gradientFill$left[local_id("A91")], NA_real_)
  expect_equal(locals$fill$gradientFill$right[local_id("A92")], 0)
  expect_equal(locals$fill$gradientFill$right[local_id("A93")], 0.5)
})
