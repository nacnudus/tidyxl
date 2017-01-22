context("Zip")

test_that("fails gracefully unzipping a missing file", {
  expect_error(zip_buffer("./examples.xlsx", "foo"), "Couldn't find 'foo' in './examples.xlsx'")
})
