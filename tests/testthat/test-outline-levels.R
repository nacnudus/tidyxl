context("Outline levels")

test_that("Outline levels are correctly retrieved", {
  cells <- xlsx_cells("outlines.xlsx")[, c("row", "col", "row_outline_level", "col_outline_level")]

  target <-
    tibble::tribble(
      ~row, ~col, ~row_outline_level, ~col_outline_level,
        2L,   2L,                  1,                  1,
        2L,   3L,                  1,                  2,
        2L,   4L,                  1,                  3,
        3L,   2L,                  3,                  1,
        3L,   3L,                  3,                  2,
        3L,   4L,                  3,                  3,
        4L,   2L,                  2,                  1,
        4L,   3L,                  2,                  2,
        4L,   4L,                  2,                  3,
        5L,   2L,                  1,                  1,
        5L,   3L,                  1,                  2,
        5L,   4L,                  1,                  3
      )
  expect_equal(cells, target)
})
