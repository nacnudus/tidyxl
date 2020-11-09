# tidyxl 1.0.7

* Update namespace in C++ code for compatibility with the latest version of
    PEGTL, wrapped in the {piton} package.

# tidyxl 1.0.6

* Fix compiler warnings.

# tidyxl 1.0.5

* Fix tests for `xlex()` (#57).

# tidyxl 1.0.4

* Compatibility: imports cell data validation rules from files created by Office
    365 (#46).

# tidyxl 1.0.3

* Noticeably faster for large files.
* Omission of blank cells with `include_blank_cells = FALSE` had a bug that
    returned blank cells as an empty row in the `xlsx_cells()` data frame.

# tidyxl 1.0.2

* Correctly constructs formulas where references are preceded by operators, e.g.
    `-F10` (#26 @cablegui).
* No longer misinterprets date formats that use underscores `_` as dates when
    the underscore is followed by a date-ish character like `M` (#24).
* Optionally omits blank cells with `include_blank_cells = FALSE` in
    `xlsx_cells()` (#25).
* Doesn't crash reading files with certain colour themes (#34 @dan-fahey).

# tidyxl 1.0.1

* Filetype checking is based on the [file
  signature](https://en.wikipedia.org/wiki/List_of_file_signatures) or "magic
  number", rather than the filename extension.  A new function `maybe_xlsx()` is
  provided for checking whether a file might be in the xlsx format.  It is
  impossible to be sure from the magic number alone, because the magic numbers
  are either common to all zip files, or common to other Microsoft Office files
  (e.g. .doc, .ppt).
* Fixed a CRAN warning.

# tidyxl 1.0.0

## New features

* `xlsx_cells()` and `xlsx_formats()` replace `tidy_xlsx()`, which has been
    deprecated.  `xlsx_cells()` returns a single data frame of all the cells in
    scope (the whole workbook, or chosen sheets), rather than a list of separate
    data frames for each sheet.  `xlsx_formats()` performs orders of magnitude
    faster.
* `xlsx_validation()` imports validation rules from cells that restrict data
    input, such as cells that require a selection from a drop-down list.  See
    the
    [vignette](https://nacnudus.github.io/tidyxl/articles/data-validation-rules.html)
    `vignette("data-validation-rules", package = "tidyxl")`.
* `xlsx_names()` imports defined names (aka named ranges/formulas), which can
    be used to filter for particular ranges of cells by name.  Use `is_range()`
    to filter for ones that are named ranges, and then read
    [joining rules to cells](https://nacnudus.github.io/tidyxl/articles/data-validation-rules.html#joining-rules-to-cells)
    for how to join cell ranges to cell addresses.  This will become easier in a
    future release.
* `is_range()` checks whether a formula is simply ranges of cells.
* `xlex()` tokenises formulas.  This is useful for detecting
    spreadsheet smells like embedded constants and deep nesting.  There is a
    [demo Shiny app](https://duncan-garmonsway.shinyapps.io/xlex/), and a
    [vignette](https://nacnudus.github.io/tidyxl/articles/smells.html)
    `vignette("smells", package = "tidyxl")`.  A vector of Excel function names
    `excel_functions` can be used to separated built-in functions from custom
    functions.  More experimental features will be implemented in the off-CRAN
    package [lexl](https://nacnudus.github.io/lexl/) before becoming part of
    tidyxl.
* `xlsx_cells()$character_formatted` is a new column for the in-cell formatting
    of text (#5).  This is for when different parts of text in a single cell
    have been formatted differently from one another.
* `is_date_format()` checks whether a number format string is a date format.
    This is useful if a cell formula contains a number formatting string (e.g.
    `TEXT(45678,"yyyy")`), and you need to know that the constant 45678 is a
    date in order to recover it at full resolution (rather than parsing the
    character output "2025" as a year).
* `xlsx_color_theme()` and it's British alias `xlsx_colour_theme()` returns the
    theme colour palette used in a file.  This is useful to monitor use of a
    corporate standard theme.
* `xlsx_color_standard` and it's British alias `xlsx_colour_standard` are data
    frames of the standard Excel palette (`red`, `blue`, etc.).
* Shared formulas are propogated to all the cells that use the same formula
    definition.  Relative cell references are handled, so that the formula
    `=A1*2` in cell `B1`  becomes `=A2*2` in cell `B2` (for more details see
    issue #7).
* Formatting of alignment and cell protection is returned (#20).

## Breaking changes and deprecations

* `tidy_xlsx()` has been deprecated in favour of `xlsx_cells()`,
    which returns a data frame of all the cells in the workbook (or in the
    requested sheets), and `xlsx_formats()`, which returns a lookup list of cell
    formats.
* In `tidy_xlsx()` and one of it's replacments `xlsx_cells()`
  * the column `content` has been replaced by `is_blank`, a logical value
    indicating whether the cell contains data.  Please replace `!is.na(content)`
    with `!is_blank` to filter out blank cells (ones with formatting but no
    value).
  * the column `formula_type` has been replaced by `is_array`, a logical value
    indicating whether the cell's formula is an array formula or not.  In Excel
    array formulas are represented visually by being surrounded by curly braces
    `{}`.
  * The order of columns has been changed so that the more useful columns are
      visible in narrow consoles.
* in `xlsx_formats()` and `tidy_xlsx()`, theme colours are given by name rather
    than by number, e.g. `"accent6"` instead of `4`.

## Minor fixes and improvements

* Certain unusual custom number formats that specify colours (e.g. `"[Cyan]0%"`)
    are no longer mis-detect as dates (#21).  `is_date_format()` tests whether a
    number format is a date format.
* `xlsx_formats()` is now thoroughly tested, and several relatively minor bugs
    fixed.  For example, `xlsx_formats(path)$local$fill$patternFill$patternType`
    consistently returns `NA` and never `"none"` when a pattern fill has not
    been set, and escape-backslashes are consistently omitted from numFmts.

## New dependency

`xlex()`, `is_range()` and the handling of relative references in shared
formulas requires a dependency on the
[piton](https://cran.r-project.org/package=piton) package, which wraps the
[PEGTL](https://github.com/taocpp/PEGTL) C++ parser generator.

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



