context("is_date_format")

test_that("is_date_format() is vectorised", {
  expect_equal(is_date_format(c("yyy/m/d h:m:s", "0%")), c(TRUE, FALSE))
})

test_that("is_date_format() is not flummoxed by colours", {
  expect_equal(is_date_format("[Black]0%"),   FALSE)
  expect_equal(is_date_format("[Blue]0%"),    FALSE)
  expect_equal(is_date_format("[Cyan]0%"),    FALSE)
  expect_equal(is_date_format("[Green]0%"),   FALSE)
  expect_equal(is_date_format("[Magenta]0%"), FALSE)
  expect_equal(is_date_format("[Red]0%"),     FALSE)
  expect_equal(is_date_format("[White]0%"),   FALSE)
  expect_equal(is_date_format("[Yellow]0%"),  FALSE)
  expect_equal(is_date_format("[black]0%"),   FALSE)
  expect_equal(is_date_format("[blue]0%"),    FALSE)
  expect_equal(is_date_format("[cyan]0%"),    FALSE)
  expect_equal(is_date_format("[green]0%"),   FALSE)
  expect_equal(is_date_format("[magenta]0%"), FALSE)
  expect_equal(is_date_format("[red]0%"),     FALSE)
  expect_equal(is_date_format("[white]0%"),   FALSE)
  expect_equal(is_date_format("[yellow]0%"),  FALSE)
})
