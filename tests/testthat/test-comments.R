context("Comments")

test_that("multiline comments are returned in full", {
  expect_equal(xlsx_cells("examples.xlsx")$comment[19],
               "commentwithtextformatting")
})

test_that("comments on blank cells are returned", {
  expect_equal(xlsx_cells("comment-on-blank-cell.xlsx")$comment,
               c(NA, "comment on blank cell"))
})
