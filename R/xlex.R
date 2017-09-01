#' @title Parse xlsx (Excel) formulas into tokens
#'
#' @description
#' \code{xlex} takes an Excel formula and separates it into tokens.  The name is
#' a bad pun on 'Excel' and 'lexer'.  It returns a dataframe, one row per token,
#' giving the token itself, its type (e.g.  'number', or 'error'), and its
#' level.
#'
#' The level is a number to show the depth of a token within nested function
#' calls.  The token 'A2' in the formula 'IF(A1=1,A2,MAX(A3,A4))' is at level 1.
#' Tokens 'A3' and 'A4' are at level 2.  The token 'IF' is at level 0, which is
#' the outermost level.
#'
#' The output isn't enough to enable computation or validation of formulas, but
#' it is enough to investigate the structure of formulas and spreadsheets.  It
#' has been tested on millions of formulas in the Enron corpus.
#'
#' @param x Character vector of length 1, giving the formula.
#'
#' @details
#' The different types of tokens are:
#'
#' \describe{
#'   \item{ref}{A cell reference/address e.g. A1 or $B2:C$14.}
#'   \item{sheet}{A sheet name, e.g. Sheet1! or 'My Sheet'!.  If the sheet is
#'     from a different file, then the file is included in this token -- usually
#'     it has been normalized to the form '[0]'.}
#'   \item{name}{A named range, or more properly a named formula.}
#'   \item{function}{An Excel or user-defined function, e.g. MAX or
#'     _xll.MY_CUSTOM_FUNCTION.}
#'   \item{error}{An error, e.g. #N/A or #REF!.}
#'   \item{bool}{TRUE or FALSE -- note that there are also functions TRUE() and
#'     FALSE().}
#'   \item{number}{All forms of numbers, e.g. 1, 1.1, -1, 1.2E3.}
#'   \item{text}{Strings inside double quotes, e.g. "Hello, World!".}
#'   \item{operator}{The usual infix operators, +, -, *, /, ^, <, <=, <>, etc.
#'     and also the range operator ':' when it is used with ranges that aren't
#'     cell addresses, e.g. INDEX(something):A1. The union operator ',' is
#'     the same symbol that is used to separate function arguments and array
#'     columns, so it is only tagged 'operator' when it is inside parentheses
#'     that are not function parentheses or array curly braces (see the
#'     examples).}
#'   \item{paren_open}{An open parenthesis '(' indicating an increase in the level
#'     of nesting, but not directly enclosing function arguments.}
#'   \item{paren_close}{As 'open', but reducing the level of nesting.}
#'   \item{open_array}{An open curly brace '\{' indicating the start of an array
#'     of constants, and an increase in the level of nesting.}
#'   \item{close_array}{As 'open_array', but ending the array of constants}
#'   \item{fun_open}{An open parenthesis '(' immediately after a function name,
#'     directly enclosing the function arguments.}
#'   \item{fun_close}{As 'fun_open' but immediately after the function
#'     arguments.}
#'   \item{separator}{A comma ',' separating function arguments or array
#'     columns, or a semicolon ';' separating array rows.}
#'   \item{DDE}{A call to a Dynamic Data Exchange server, usually normalized to
#'     the form [1]!'DDE_parameter=1', but the full form is
#'     'ABCD'|'EFGH'!'IJKL'.}
#'   \item{space}{Some old files haven't stripped formulas of meaningless
#'     spaces. They are returned as 'space' tokens so that the original formula
#'     can always be reconstructed by concatenating all tokens.}
#'   \item{other}{If you see this, then something has gone wrong -- please
#'     report it at https://github.com/nacnudus/tidyxl/issues with a
#'     reproducible example (e.g. using the reprex package).}
#' }
#'
#' Every part of the original formula is returned as a token, so the original
#' formula can be reconstructed by concatenating the tokens.  If that doesn't
#' work, please report it at https://github.com/nacnudus/tidyxl/issues with a
#' reproducible example (e.g. using the reprex package).
#'
#' The XLParser project was a great help in creating the grammar.
#' \url{https://github.com/spreadsheetlab/XLParser}.
#'
#' @return
#' A data frame (a tibble, if you use the tidyverse) one row per token,
#' giving the token itself, its type (e.g.  'number', or 'error'), and its
#' level.
#'
#' @export
#' @examples
#' # All explicit cell references/addresses are returned as a single 'ref' token.
#' xlex("A1")
#' xlex("A$1")
#' xlex("$A1")
#' xlex("$A$1")
#' xlex("A1:B2")
#' xlex("1:1") # Whole row
#' xlex("A:B") # Whole column
#'
#' # If one part of an address is a name or a function, then the colon ':' is
#' # regarded as a 'range operator', so is tagged 'operator'.
#' xlex("A1:SOME.NAME")
#' xlex("SOME_FUNCTION():B2")
#' xlex("SOME_FUNCTION():SOME.NAME")
#'
#' # Sheet names are recognised by the terminal exclamation mark '!'.
#' xlex("Sheet1!A1")
#' xlex("'Sheet 1'!A1")       # Quoted names may contain some punctuation
#' xlex("'It''s a sheet'!A1") # Quotes are escaped by doubling
#'
#' # Sheets can be ranged together in so-called 'three-dimensional formulas'.
#' # Both sheets are returned in a single 'sheet' token.
#' xlex("Sheet1:Sheet2!A1")
#' xlex("'Sheet 1:Sheet 2'!A1") # Quotes surround both (rather than each) sheet
#'
#' # Sheets from other files are prefixed by the filename, which Excel
#' # normalizes the filenames into indexes.  Either way, xlex() includes the
#' # file/index in the 'sheet' token.
#' xlex("[1]Sheet1!A1")
#' xlex("'[1]Sheet 1'!A1") # Quotes surround both the file index and the sheet
#' xlex("'C:\\My Documents\\[file.xlsx]Sheet1'!A1")
#'
#' # Function names are recognised by the terminal open-parenthesis '('.  There
#' # is no distinction between custom functions and built-in Excel functions.
#' # The open-parenthesis is tagged 'fun_open', and the corresponding
#' # close-parenthesis at the end of the arguments is tagged 'fun_close'.
#' xlex("MAX(1,2)")
#' xlex("_xll.MY_CUSTOM_FUNCTION()")
#'
#' # Named ranges (properly called 'named formulas') are a last resort after
#' # attempting to match a function (ending in an open parenthesis '(') or a
#' # sheet (ending in an exclamation mark '!')
#' xlex("MY_NAMED_RANGE")
#'
#' # Some cell addresses/references, functions and names can look alike, but
#' # xlex() should always make the right choice.
#' xlex("XFD1")     # A cell in the maximum column in Excel
#' xlex("XFE1")     # Beyond the maximum column, must be a named range/formula
#' xlex("A1048576") # A cell in the maximum row in Excel
#' xlex("A1048577") # Beyond the maximum row, must be a named range/formula
#' xlex("LOG10")    # A cell address
#' xlex("LOG10()")  # A log function
#' xlex("LOG:LOG")  # The whole column 'LOG'
#' xlex("LOG")      # Not a cell address, must be a named range/formula
#' xlex("LOG()")    # Another log function
#' xlex("A1.2!A1")  # A sheet called 'A1.2'
#'
#' # Text is surrounded by double-quotes.
#' xlex("\"Some text\"")
#' xlex("\"Some \"\"text\"\"\"") # Double-quotes within text are escaped by
#'
#' # Numbers are signed where it makes sense, and can be scientific
#' xlex("1")
#' xlex("1.2")
#' xlex("-1")
#' xlex("-1-1")
#' xlex("-1+-1")
#' xlex("MAX(-1-1)")
#' xlex("-1.2E-3")
#'
#' # Booleans can be constants or functions, and names can look like booleans,
#' # but xlex() should always make the right choice.
#' xlex("TRUE")
#' xlex("TRUEISH")
#' xlex("TRUE!A1")
#' xlex("TRUE()")
#'
#' # Errors are tagged 'error'
#' xlex("#DIV/0!")
#' xlex("#N/A")
#' xlex("#NAME?")
#' xlex("#NULL!")
#' xlex("#NUM!")
#' xlex("#REF!")
#' xlex("#VALUE!")
#'
#' # Operators with more than one character are treated as single tokens
#' xlex("1<>2")
#' xlex("1<=2")
#' xlex("1<2")
#' xlex("1=2")
#' xlex("1&2")
#' xlex("1%")   # postfix operator
#'
#' # The union operator is a comma ',', which is the same symbol that is used
#' # to separate function arguments or array columns.  It is tagged 'operator'
#' # only when it is inside parentheses that are not function parentheses or
#' # array curly braces.  The curly braces are tagged 'array_open' and
#' # 'array_close'.
#' tidyxl::xlex("A1,B2") # invalid formula, defaults to 'union' to avoid a crash
#' tidyxl::xlex("(A1,B2)")
#' tidyxl::xlex("MAX(A1,B2)")
#' tidyxl::xlex("SMALL((A1,B2),1)")
#'
#' # Function arguments are separated by commas ',', which are tagged
#' # 'separator'.
#' xlex("MAX(1,2)")
#'
#' # Nested functions are marked by an increase in the 'level'.  The level
#' # increases inside parentheses, rather than at the parentheses.  Curly
#' # braces, for arrays, have the same behaviour, as do subexpressions inside
#' # ordinary parenthesis, tagged 'paren_open' and 'paren_close'.
#' xlex("MAX(MIN(1,2),3)")
#' xlex("{1,2;3,4}")
#' xlex("1*(2+3)")
#'
#' # Arrays are marked by opening and closing curly braces, with comma ','
#' # between columns, and semicolons ';' between rows  Commas and semicolons are
#' # both tagged 'separator'.  Arrays contain only constants, which are
#' # booleans, numbers, text, and errors.
#' xlex("MAX({1,2;3,4})")
#' xlex("=MAX({-1E-2,TRUE;#N/A,\"Hello, World!\"})")
#'
#' # Structured references are surrounded by square brackets.  Subexpressions
#' # may also be surrounded by square brackets, but xlex() returns the whole
#' # expression in a single 'structured_ref' token.
#' xlex("[@col2]")
#' xlex("SUM([col22])")
#' xlex("Table1[col1]")
#' xlex("Table1[[col1]:[col2]]")
#' xlex("Table1[#Headers]")
#' xlex("Table1[[#Headers],[col1]]")
#' xlex("Table1[[#Headers],[col1]:[col2]]")
#'
#' # DDE calls (Dynamic Data Exchange) are normalized by Excel into indexes.
#' # Either way, xlex() includes everything in one token.
#' xlex("[1]!'DDE_parameter=1'")
#' xlex("'Quote'|'NYSE'!ZAXX")
#' # Meaningless spaces that appear in some old files are returned as 'space'
#' # tokens, so that the original formula can still be recovered by
#' # concatenating all the tokens.  Spaces between function names and their open
#' # parenthesis have not been observed, so are not permitted.
#' xlex(" MAX( A1 ) ")
xlex <- function(x) {
  if (length(x) != 1) {
    stop("'x' must be a character vector of length 1")
  }
  if (!is.character(x)) {
    stop("'x' must be a character vector of length 1")
  }
  xlex_(x)
}
