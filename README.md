---
output: github_document
---

<!-- README.md is generated from README.Rmd. Please edit that file -->



***This package is in early development.***

# tidyxl

Imports untidy Excel files into R without forcing the data into data frames.
Exposes cell content, position and formatting in a tidy structure for further
manipulation.  Supports '.xlsx' and '.xlsm' via the embedded
[RapidXML](http://rapidxml.sourceforge.net) C++ library.  Does not support
'.xlsb' or '.xls'. 

## Introduction

Information in in many spreadsheets cannot be easily imported into R.  Why?

Most R packages that import spreadsheets have difficulty unless the layout of
the spreadsheet conforms to a strict definition of a 'table', e.g.:

* observations in rows
* variables in columns
* a single header row
* all information represented by characters, whether textual, logical, or
  numeric

These rules are designed to eliminate ambiguity in the interpretation of the
information.  But most spreadsheeting software relaxes these rules in a trade of
ambiguity for expression via other media:

* proximity (other than headers, i.e. other than being the first value at the
  top of a column)
* formatting (colours and borders)

Humans can usually resolve the ambiguities with contextual knowledge, but
computers are limited by their ignorance.  Programmers are hampered by:

* their language's expressiveness
* loss of information in transfer from spreadsheet to programming library

Information is lost when software discards it in order to force the data into
tabular form.  Sometimes date formatting is retained, but mostly formatting 
is lost, and position has to be inferred again.

The `tidyxl` package addresses the programmer's problems by not discarding
information.  It imports the content, position and formatting of cells, leaving
it up to the user to associate the different forms of information, and to
re-encode them in tabular form without loss.  Further packages are planned to
assist with that step.
