context("Comments")

test_that("multiline comments are returned in full", {
  expect_equal(tidy_xlsx("examples.xlsx")$data$Sheet1$comment[19],
               "commentwithtextformatting")
})

test_that("comments on blank cells are returned", {
  expected <-
    structure(list(address = c("A2", "A1"),
                   row = c(2L, 1L),
                   col = c(1L, 1L),
                   content = c("0", NA),
                   formula = c(NA_character_, NA_character_),
                   formula_type = c(NA_character_, NA_character_),
                   formula_ref = c(NA_character_, NA_character_),
                   formula_group = c(NA_integer_, NA_integer_),
                   type = c("s", NA),
                   data_type = c("character", "blank"),
                   error = c(NA_character_, NA_character_),
                   logical = c(NA, NA),
                   numeric = c(NA_real_, NA_real_),
                   date = c(NA_real_, NA_real_),
                   character = c("non-blank cell", NA),
                   comment = c(NA, "comment on blank cell"),
                   height = c(22.5, 22.5),
                   width = c(11.375, 11.375),
                   style_format = c("Normal", "Normal"),
                   local_format_id = c(1L, 1L)),
              .Names = c("address",
                         "row",
                         "col",
                         "content",
                         "formula",
                         "formula_type",
                         "formula_ref",
                         "formula_group",
                         "type",
                         "data_type",
                         "error",
                         "logical",
                         "numeric",
                         "date",
                         "character",
                         "comment",
                         "height",
                         "width",
                         "style_format",
                         "local_format_id"),
              class = c("tbl_df", "tbl", "data.frame"),
              row.names = c(NA, -2L))
  expect_equal(tidy_xlsx("comment-on-blank-cell.xlsx")$data$Sheet1, expected)
})
