# tidyxl 0.2.3.9000

* Tokenizes formulas with `xlex()` and plots the parse tree with `plot_xlex()`.
    There's a demo Shiny app `demo_xlex()`, hosted on [shinyapps.io](), and a
    new vignette, `vignette("smells", package = "tidyxl")`, showing how to use
    `xlex()` to detect spreadsheet smells like embedded constants and deep
    nesting.  There are also some intermediate functions for working with the
    parse tree: `xlex_edges()`, `xlex_vertices`, and `xlex_igraph()`.
* Propogates shared formulas, handling relative cell references (for details see
    issue #7).  This introduces a dependency on
    [piton](https://cran.r-project.org/web/packages/piton/index.html), which
    wraps the [PEGTL](https://github.com/taocpp/PEGTL) parser generator.

# tidyxl 0.2.3

* Imports dates on or before 1900-02-28 dates correctly, and only warns on the impossible date 1900-02-29, returning NA (following [readxl](https://github.com/tidyverse/readxl/commit/c9a54ae9ce0394808f6d22e8ef1a7a647b2d92bb)).
* Fixes subsecond rounding (following [readxl](https://github.com/tidyverse/readxl/commit/63ef215f57322dd5d7a27799a2a3fe463bd39fc7) (fixes #14))
* Imports styles correctly from LibreOffice files (interprets 'true' and 'false'
  as booleans, as well as the 0 and 1 used by Microsoft Excel, and defaults to
  'true' when not present, e.g. applyFill)
* Fixes a bug that caused some LibreOffice files to crash R, when styles were
    declared with gaps in the sequence of xfIds.
* Imports comments attached to blank cells (fixes #10)
* Includes the sheet and cell address in warnings of impossible datetimes from
    1900-02-29.

# tidyxl 0.2.1.9000

* Checks the value of the `date1904` attribute for `"false"` or `"1"` to support files
  created by the `openxlsx` package (#8).
* Fixed a bug that only imported the first line of multiline comments (#9).
* Encodes cell and comment text as UTF-8 (#12).
* Finds worksheets more reliably in files not created by Excel (part of #13).
* Falls back to default styles when none defined (#13).
* Imports dates with greater precision (part of #14).
* Fixed the order of worksheets (#15)

# tidyxl 0.2.1

* Fixed a major bug: dates were parsed incorrectly because the offsets for the
  1900 and 1904 systems were the wrong way around.
* Added a warning when dates suffer from the Excel off-by-one bug.
* Fixed a bug that misinterpreted some number formats as dates (confused by the
  'd' in '[Red]', which looks like a 'd' in 'd/m/y'.)
* Added support for xlsx files created by Gnumeric (a single, unnamed cell
  formatting style).
* Fixed the checkUserInterrupt to work every 1000th cell.
* Added many tests.
* Removed lots of redundant code.

# tidyxl 0.2.0

* Breaking change: Look up style formats by style name (e.g.
  `"Normal"`) instead of by index integer.  All the vectors under
  `x$formats$style` are named according to the style names.
  `x$data$sheet$style_format_id` has been renamed to `x$data$sheet$style_format`
  and its type changed from integer (index into style formats) to character
  (still an index, but looking up the named vectors by name).  There are
  examples in the README and vignette.
* Simplified some variable names in the C++ code.

# tidyxl 0.1.1

* Fixed many compilation errors on Travis/AppVeyor/r-hub
* Many small non-breaking changes

# tidyxl 0.1.0

* Added a `NEWS.md` file to track changes to the package.



