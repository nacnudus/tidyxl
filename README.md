
<!-- README.md is generated from README.Rmd. Please edit that file -->
tidyxl
======

[![Travis-CI Build Status](https://travis-ci.org/nacnudus/tidyxl.svg?branch=master)](https://travis-ci.org/nacnudus/tidyxl) [![AppVeyor Build Status](https://ci.appveyor.com/api/projects/status/github/nacnudus/tidyxl?branch=master&svg=true)](https://ci.appveyor.com/project/nacnudus/tidyxl) [![Cran Status](http://www.r-pkg.org/badges/version/tidyxl)](https://cran.r-project.org/web/packages/tidyxl/index.html) [![Cran Downloads](https://cranlogs.r-pkg.org/badges/tidyxl)](https://www.r-pkg.org/pkg/tidyxl) [![codecov](https://codecov.io/github/nacnudus/tidyxl/coverage.svg?branch=master)](https://codecov.io/gh/nacnudus/tidyxl)

[tidyxl](https://github.com/nacnudus/tidyxl) imports non-tabular data from Excel files into R. It exposes cell content, position, formatting and comments in a tidy structure for further manipulation, especially by the [unpivotr](https://github.com/nacnudus/unpivotr) package. It supports the xml-based file formats '.xlsx' and '.xlsm' via the embedded [RapidXML](http://rapidxml.sourceforge.net) C++ library. It does not support the binary file formats '.xlsb' or '.xls'.

Mailing list
------------

For bugs and/or issues, create a new issue on [GitHub](https://github.com/nacnudus/tidyxl/issues) For other questions or comments, please subscribe to the [tidyxl-devel mailing list](https://groups.google.com/forum/#!forum/tidyxl-devel). You must be a member to post messages, but anyone can read the archived discussions.

Installation
------------

``` r
devtools::install_github("nacnudus/tidyxl")
```

Examples
--------

The package includes a spreadsheet, 'titanic.xlsx', which contains the following pivot table:

``` r
ftable(Titanic, row.vars = 1:2)
#>              Age      Child     Adult    
#>              Survived    No Yes    No Yes
#> Class Sex                                
#> 1st   Male                0   5   118  57
#>       Female              0   1     4 140
#> 2nd   Male                0  11   154  14
#>       Female              0  13    13  80
#> 3rd   Male               35  13   387  75
#>       Female             17  14    89  76
#> Crew  Male                0   0   670 192
#>       Female              0   0     3  20
```

The multi-row column headers make this difficult to import. A popular package for importing spreadsheets coerces the pivot table into a dataframe. It treats the second header row as though it were observations.

``` r
titanic <- system.file("extdata/titanic.xlsx", package = "tidyxl")
readxl::read_excel(titanic)
#> # A tibble: 10 x 7
#>     X__1   X__2      Age Child  X__3 Adult  X__4
#>    <chr>  <chr>    <chr> <chr> <chr> <chr> <chr>
#>  1  <NA>   <NA> Survived    No   Yes    No   Yes
#>  2 Class    Sex     <NA>  <NA>  <NA>  <NA>  <NA>
#>  3   1st   Male     <NA>     0     5   118    57
#>  4  <NA> Female     <NA>     0     1     4   140
#>  5   2nd   Male     <NA>     0    11   154    14
#>  6  <NA> Female     <NA>     0    13    13    80
#>  7   3rd   Male     <NA>    35    13   387    75
#>  8  <NA> Female     <NA>    17    14    89    76
#>  9  Crew   Male     <NA>     0     0   670   192
#> 10  <NA> Female     <NA>     0     0     3    20
```

[tidyxl](https://github.com/nacnudus/tidyxl) doesn't coerce the pivot table into a data frame. Instead, it represents each cell in its own row, where it describes the cell's address, value and other properties.

``` r
library(tidyxl)
x <- tidy_xlsx(titanic)$data$Sheet1
# Specific sheets can be requested using `tidy_xlsx(file, sheet)`
str(x)
#> Classes 'tbl_df', 'tbl' and 'data.frame':    60 obs. of  20 variables:
#>  $ address        : chr  "C1" "D1" "E1" "F1" ...
#>  $ row            : int  1 1 1 1 1 2 2 2 2 2 ...
#>  $ col            : int  3 4 5 6 7 3 4 5 6 7 ...
#>  $ content        : chr  "0" "1" NA "2" ...
#>  $ formula        : chr  NA NA NA NA ...
#>  $ formula_type   : chr  NA NA NA NA ...
#>  $ formula_ref    : chr  NA NA NA NA ...
#>  $ formula_group  : int  NA NA NA NA NA NA NA NA NA NA ...
#>  $ type           : chr  "s" "s" NA "s" ...
#>  $ data_type      : chr  "character" "character" "blank" "character" ...
#>  $ error          : chr  NA NA NA NA ...
#>  $ logical        : logi  NA NA NA NA NA NA ...
#>  $ numeric        : num  NA NA NA NA NA NA NA NA NA NA ...
#>  $ date           : POSIXct, format: NA NA ...
#>  $ character      : chr  "Age" "Child" NA "Adult" ...
#>  $ comment        : chr  NA NA NA NA ...
#>  $ height         : num  15 15 15 15 15 15 15 15 15 15 ...
#>  $ width          : num  8.38 8.38 8.38 8.38 8.38 8.38 8.38 8.38 8.38 8.38 ...
#>  $ style_format   : chr  "Normal" "Normal" "Normal" "Normal" ...
#>  $ local_format_id: int  2 3 3 3 3 2 3 3 3 3 ...
```

In this structure, the cells can be found by filtering.

``` r
x[x$data_type == "character", c("address", "character")]
#> # A tibble: 22 x 2
#>    address character
#>      <chr>     <chr>
#>  1      C1       Age
#>  2      D1     Child
#>  3      F1     Adult
#>  4      C2  Survived
#>  5      D2        No
#>  6      E2       Yes
#>  7      F2        No
#>  8      G2       Yes
#>  9      A3     Class
#> 10      B3       Sex
#> # ... with 12 more rows
x[x$row == 4, c("address", "character", "numeric")]
#> # A tibble: 6 x 3
#>   address character numeric
#>     <chr>     <chr>   <dbl>
#> 1      A4       1st      NA
#> 2      B4      Male      NA
#> 3      D4      <NA>       0
#> 4      E4      <NA>       5
#> 5      F4      <NA>     118
#> 6      G4      <NA>      57
```

### Formatting

The original spreadsheet has formatting applied to the cells. This can also be retrieved using [tidyxl](https://github.com/nacnudus/tidyxl).

![](./vignettes/titanic-screenshot.png)

Formatting is available by using the columns `local_format_id` and `style_format` as indexes into a separate list-of-lists structure. 'Local' formatting is the most common kind, applied to individual cells. 'Style' formatting is usually applied to blocks of cells, and defines several formats at once. Here is a screenshot of the styles buttons in Excel.

![](./vignettes/styles-screenshot.png)

Formatting can be looked up as follows.

``` r
# Bold
formats <- tidy_xlsx(titanic)$formats
formats$local$font$bold
#> [1] FALSE  TRUE FALSE FALSE
x[x$local_format_id %in% which(formats$local$font$bold),
  c("address", "character")]
#> # A tibble: 4 x 2
#>   address character
#>     <chr>     <chr>
#> 1      C1       Age
#> 2      C2  Survived
#> 3      A3     Class
#> 4      B3       Sex

# Yellow fill
formats$local$fill$patternFill$fgColor$rgb
#> [1] NA         NA         NA         "FFFFFF00"
x[x$local_format_id %in%
  which(formats$local$fill$patternFill$fgColor$rgb == "FFFFFF00"),
  c("address", "numeric")]
#> # A tibble: 2 x 2
#>   address numeric
#>     <chr>   <dbl>
#> 1     F11       3
#> 2     G11      20

# Styles by name
formats$style$font$name["Normal"]
#>    Normal 
#> "Calibri"
head(x[x$style_format == "Normal", c("address", "character")])
#> # A tibble: 6 x 2
#>   address character
#>     <chr>     <chr>
#> 1      C1       Age
#> 2      D1     Child
#> 3      E1      <NA>
#> 4      F1     Adult
#> 5      G1      <NA>
#> 6      C2  Survived
```

To see all the available kinds of formats, use `str(formats)`.

### Comments

Comments are available alongside cell values.

``` r
x[!is.na(x$comment), c("address", "comment")]
#> # A tibble: 1 x 2
#>   address                                                     comment
#>     <chr>                                                       <chr>
#> 1     G11 All women in the crew worked in the victualling department.
```

### Formulas

Formulas are available, but with a few quirks.

``` r
options(width = 120)
y <- tidy_xlsx(system.file("/extdata/examples.xlsx", package = "tidyxl"),
               "Sheet1")$data[[1]]
y[!is.na(y$formula),
  c("address", "formula", "formula_type", "formula_ref", "formula_group",
    "error", "logical", "numeric", "date", "character")]
#> # A tibble: 15 x 10
#>    address              formula formula_type formula_ref formula_group   error logical numeric       date     character
#>      <chr>                <chr>        <chr>       <chr>         <int>   <chr>   <lgl>   <dbl>     <dttm>         <chr>
#>  1      A1                  1/0         <NA>        <NA>            NA #DIV/0!      NA      NA         NA          <NA>
#>  2     A14                  1=1         <NA>        <NA>            NA    <NA>    TRUE      NA         NA          <NA>
#>  3     A15                 A4+1         <NA>        <NA>            NA    <NA>      NA    1338         NA          <NA>
#>  4     A16      DATE(2017,1,18)         <NA>        <NA>            NA    <NA>      NA      NA 2017-01-18          <NA>
#>  5     A17  "\"Hello, World!\""         <NA>        <NA>            NA    <NA>      NA      NA         NA Hello, World!
#>  6     A19                A18+1         <NA>        <NA>            NA    <NA>      NA       2         NA          <NA>
#>  7     B19                A18+2         <NA>        <NA>            NA    <NA>      NA       3         NA          <NA>
#>  8     A20                A19+1       shared     A20:A21             0    <NA>      NA       3         NA          <NA>
#>  9     B20                A19+2       shared     B20:B21             1    <NA>      NA       4         NA          <NA>
#> 10     A21                            shared        <NA>             0    <NA>      NA       4         NA          <NA>
#> 11     B21                            shared        <NA>             1    <NA>      NA       5         NA          <NA>
#> 12     A22 SUM(A19:A21*B19:B21)        array         A22            NA    <NA>      NA      38         NA          <NA>
#> 13     A23      A19:A20*B19:B20        array     A23:A24            NA    <NA>      NA       6         NA          <NA>
#> 14     A25       [1]Sheet1!$A$1         <NA>        <NA>            NA    <NA>      NA      NA         NA        normal
#> 15     A94               50*10%         <NA>        <NA>            NA    <NA>      NA       5         NA          <NA>
```

The top five cells show that the results of formulas are available as usual in the columns `error`, `logical`, `numeric`, `date`, and `character`.

Cells `A20` and `A21` share a formula definition. The formula is given against cell `A20`, and assigned to `formula_group` `0`, which spans the cells given by the `formula_ref`, A20:A21. A spreadsheet application would infer that cell `A21` had the formula `A20+1`. Cells `B20` and `B21` are similar. The roadmap for [tidyxl](https://github.com/nacnudus/tidyxl) includes de-normalising shared formulas. If you can suggest how to tokenize Excel formulas, then please contact me.

Cell `A22` contains an array formula, which, in a spreadsheet application, would appear with curly braces `{SUM(A19:A21*B19:B21)}`. Cells `A23` and `A24` contain a single multi-cell array formula (single formula, multi-cell result), indicated by the `formula_ref`, but unlike cells `A20:A21` and `B20:B21`, the `formula` for A24 is NA rather than blank (`""`), and it doesn't have a `formula_group`.

Cell `A25` contains a formula that refers to another file. The `[1]` is an index into a table of files. The roadmap [tidyxl](https://github.com/nacnudus/tidyxl) for tidyxl includes de-referencing such numbers.

[tidyxl](https://github.com/nacnudus/tidyxl) imports the same table into a format suitable for non-tabular processing (see e.g. the [unpivotr](https://github.com/nacnudus/unpivotr) package in 'Similar projects' below).

Philosophy
----------

Information in in many spreadsheets cannot be easily imported into R. Why?

Most R packages that import spreadsheets have difficulty unless the layout of the spreadsheet conforms to a strict definition of a 'table', e.g.:

-   observations in rows
-   variables in columns
-   a single header row
-   all information represented by characters, whether textual, logical, or numeric

These rules are designed to eliminate ambiguity in the interpretation of the information. But most spreadsheeting software relaxes these rules in a trade of ambiguity for expression via other media:

-   proximity (other than headers, i.e. other than being the first value at the top of a column)
-   formatting (colours and borders)

Humans can usually resolve the ambiguities with contextual knowledge, but computers are limited by their ignorance. Programmers are hampered by:

-   their language's expressiveness
-   loss of information in transfer from spreadsheet to programming library

Information is lost when software discards it in order to force the data into tabular form. Sometimes date formatting is retained, but mostly formatting is lost, and position has to be inferred again.

[tidyxl](https://github.com/nacnudus/tidyxl) addresses the programmer's problems by not discarding information. It imports the content, position and formatting of cells, leaving it up to the user to associate the different forms of information, and to re-encode them in tabular form without loss. The [unpivotr](https://github.com/nacnudus/unpivotr) package has been developed to assist with that step.

Similar projects
----------------

[tidyxl](https://github.com/nacnudus/tidyxl) was originally derived from [readxl](https://github.com/hadley/readxl) and still contains some of the same code, hence it inherits the GPL-3 licence. [readxl](https://github.com/hadley/readxl) is intended for importing tabular data with a single row of column headers, whereas [tidyxl](https://github.com/nacnudus/tidyxl) is more general, and less magic.

The [rsheets](https://github.com/rsheets) project of several R packages is in the early stages of importing spreadsheet information from Excel and Google Sheets into R, manipulating it, and potentially parsing and processing formulas and writing out to spreadsheet files. In particular, [jailbreaker](https://github.com/rsheets/jailbreakr) attempts to extract non-tabular data from spreadsheets into tabular structures automatically via some clever algorithms.

[tidyxl](https://github.com/nacnudus/tidyxl) differs from [rsheets](https://github.com/rsheets) in scope ([tidyxl](https://github.com/nacnudus/tidyxl) will never import charts, for example), and implementation ([tidyxl](https://github.com/nacnudus/tidyxl) is implemented mainly in C++ and is quite fast, only a little slower than [readxl](https://github.com/hadley/readxl)). [unpivotr](https://github.com/nacnudus/unpivotr) is a package related to [tidyxl](https://github.com/nacnudus/tidyxl) that provides tools for unpivoting complex and non-tabular data layouts using I not AI (intelligence, not artificial intelligence). In this way it corresponds to [jailbreaker](https://github.com/rsheets/jailbreakr), but with a different philosophy.

Roadmap
-------

-   \[ \] Parse shared formulas and propagate to all associated cells.
-   \[ \] Propagate array formulas to all associated cells.
-   \[x\] Parse dates
-   \[x\] Detect cell types (date, boolean, string, number)
-   \[x\] Implement formatting import in C++ for speed.
-   \[x\] Write more tests
