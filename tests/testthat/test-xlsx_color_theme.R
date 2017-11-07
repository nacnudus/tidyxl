context("xlsx_color_theme")

library(tibble)

test_that("xlsx_color_theme() and its British alias work", {
  expect_equal(xlsx_color_theme("./examples.xlsx"),
               tribble(      ~theme,       ~rgb,
                      "background1", "FFFFFFFF",
                            "text1", "FF000000",
                      "background2", "FFEEECE1",
                            "text2", "FF1F497D",
                          "accent1", "FF4F81BD",
                          "accent2", "FFC0504D",
                          "accent3", "FF9BBB59",
                          "accent4", "FF8064A2",
                          "accent5", "FF4BACC6",
                          "accent6", "FFF79646",
                        "hyperlink", "FF0000FF",
               "followed-hyperlink", "FF800080"))
})
