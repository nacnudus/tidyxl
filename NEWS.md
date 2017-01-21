# tidyxl 0.1.1.9000

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



