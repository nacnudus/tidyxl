# vdiffr::manage_cases()

context("plot_xlex()")

if (Sys.info()["sysname"] == "Linux") {
  test_that("plot_xlex() looks correct", {
    vdiffr::expect_doppelganger("plot_xlex() looks about right",
      plot_xlex(xlex("MIN(3,MAX(2,A1))")))
  })
}

x <- xlex("MIN(3,MAX(2,A1))")

test_that("xlex_edges() is correct", {
  expect_equal(
    xlex_edges(x),
    data.frame(from = c("0", "0",  "0", "2", "2", "2", "2",  "2", "6", "6", "6"),
               to   = c("1", "2", "11", "3", "4", "5", "6", "10", "7", "8", "9"),
               stringsAsFactors = FALSE))
})

test_that("xlex_vertices() is correct", {
  expect_equal(xlex_vertices(x),
    data.frame(id = c("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11"),
               label = c("root", "MIN\nfunction", "(\nfun_open", "3\nnumber",
                         ",\nseparator", "MAX\nfunction", "(\nfun_open",
                         "2\nnumber", ",\nseparator", "A1\nref", ")\nfun_close",
                         ")\nfun_close"),
               type = c("root", "function", "fun_open", "number", "separator",
                        "function", "fun_open", "number", "separator", "ref",
                        "fun_close", "fun_close"),
               token = c("root", "MIN", "(", "3", ",", "MAX", "(", "2", ",", "A1", ")", ")"),
               stringsAsFactors = FALSE))
})

test_that("xlex_igraph() doesn't error", {
  # Can't test equality because can't construct an igraph with deparse
  expect_error(xlex_igraph(x), NA)
})
