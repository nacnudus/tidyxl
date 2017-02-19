# tidyxl 0.2.1.9000

* Checks the value of the `date1904` attribute for `"false"` or `"1"` to support files
  created by the `openxlsx` package (#8).
* Fixed a bug that only imported the first line of multiline comments (#9).
* Encodes cell and comment text as UTF-8 (#12).
* Finds worksheets more reliably in files not created by Excel (part of #13).
* Imports dates with greater precision (part of #14).

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



