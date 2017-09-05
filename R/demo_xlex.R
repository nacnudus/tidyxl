#' @title Run a shiny app to demonstrate xlex()
#'
#' @description Requires the `shiny`, `igraph` and `ggraph` packages, which you
#' can install with `install.packages(c("shiny", "igraph", "ggraph"))`.  Runs a
#' shiny app to demonstrate tokenizing an Excel formula with `xlex()` and
#' plotting the parse tree with `plot_xlex()`
#'
#' @seealso [tidyxl::xlex()], [tidyxl::plot_xlex()]
#' @export
#' @examples
#' \dontrun{
#'   demo_xlex()
#' }
demo_xlex <- function() {
  if (!requireNamespace("shiny", quietly = TRUE)
      && requireNamespace("igraph", quietly = TRUE)
      && requireNamespace("ggraph", quietly = TRUE)) {
    stop("demo_xlex() requires the packages 'shiny', 'igraph', and 'ggraph'.",
         call. = FALSE)
  } else {
    shiny::shinyApp(
      ui = shiny::fluidPage(
        shiny::titlePanel(shiny::HTML("Demo of <b><a href=\"https://nacnudus.github.io/tidyxl/\">tidyxl</a></b>::<b><a href=\"https://nacnudus.github.io/tidyxl/articles/smells.html\">xlex()</a></b>: Tokenize Excel formulas with R")),
        shiny::sidebarLayout(
          shiny::sidebarPanel(
            width = 3,
            shiny::textAreaInput("formula",
                                 "Formula",
                                 "MIN(3,MAX(2,A1))",
                                 resize = "vertical"),
            shiny::htmlOutput("call"),
            shiny::tableOutput("tree")
          ),
          shiny::mainPanel(
            shiny::plotOutput("plot")
          )
        )
      ),
      server = function(input, output) {
        parse_tree <- shiny::reactive({
          xlex(input$formula)
        })
        output$call <- shiny::renderText(paste0("<pre>library(tidyxl)\n",
                                                "parse_tree <- xlex(\"",
                                                  input$formula,
                                                "\")\n",
                                                "parse_tree\n",
                                                "plot_xlex(parse_tree)"))
        output$tree <- shiny::renderTable({parse_tree()})
        output$plot <- shiny::renderPlot({
          plot_xlex(xlex(input$formula))
        })
      })
  }
}
