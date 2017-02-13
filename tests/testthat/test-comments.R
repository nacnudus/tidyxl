context("Comments")

test_that("multiline comments are returned in full", {
  expect_equal(tidy_xlsx("examples.xlsx")$data$Sheet1$comment[19],
               "commentwithtextformatting")
})
