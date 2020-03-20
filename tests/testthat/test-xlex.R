context("xlex()")

library(tibble)

# Wrapper of xlex() that doesn't have the `"xlex"` class, so can be compared
# with a simple tribble.
un_xlex <- function(...) {
  x <- xlex(...)
  class(x) <- class(x)[-which(class(x) == "xlex")]
  x
}

test_that("All explicit cell references/addresses are returned as a single 'ref' token", {
  expect_equal(un_xlex("A1"),
               tribble(~level, ~type,  ~token,
                           0L, "ref",    "A1"))
  expect_equal(un_xlex("A$1"),
               tribble(~level, ~type,  ~token,
                           0L, "ref",   "A$1"))
  expect_equal(un_xlex("$A1"),
               tribble(~level, ~type,  ~token,
                           0L, "ref",   "$A1"))
  expect_equal(un_xlex("$A$1"),
               tribble(~level, ~type,  ~token,
                           0L, "ref",  "$A$1"))
  expect_equal(un_xlex("A1:B2"),
               tribble(~level, ~type,  ~token,
                           0L, "ref", "A1:B2"))
  expect_equal(un_xlex("1:1"),
               tribble(~level, ~type,  ~token,
                           0L, "ref",   "1:1"))
  expect_equal(un_xlex("A:B"),
               tribble(~level, ~type,  ~token,
                           0L, "ref",   "A:B"))
})

test_that("colon is tagged 'operator' between range and name/function", {
  expect_equal(un_xlex("A1:SOME.NAME"),
               tribble(~level,      ~type,      ~token,
                           0L,      "ref",        "A1",
                           0L, "operator",         ":",
                           0L,     "name", "SOME.NAME"))
  expect_equal(un_xlex("SOME_FUNCTION():B2"),
               tribble(~level,       ~type,          ~token,
                           0L,  "function", "SOME_FUNCTION",
                           0L,  "fun_open",             "(",
                           0L, "fun_close",             ")",
                           0L,  "operator",             ":",
                           0L,       "ref",            "B2"))
  expect_equal(un_xlex("SOME_FUNCTION():SOME.NAME"),
               tribble(~level,       ~type,          ~token,
                           0L,  "function", "SOME_FUNCTION",
                           0L,  "fun_open",             "(",
                           0L, "fun_close",             ")",
                           0L,  "operator",             ":",
                           0L,      "name",   "SOME.NAME"))
})

test_that("colon is tagged 'operator' between range and name/function", {
  expect_equal(un_xlex("A1:SOME.NAME"),
               tribble(~level,      ~type,        ~token,
                           0L,      "ref",          "A1",
                           0L, "operator",           ":",
                           0L,     "name", "SOME.NAME"))
  expect_equal(un_xlex("SOME_FUNCTION():B2"),
               tribble(~level,       ~type,          ~token,
                           0L,  "function", "SOME_FUNCTION",
                           0L,  "fun_open",             "(",
                           0L, "fun_close",             ")",
                           0L,  "operator",             ":",
                           0L,       "ref",          "B2"))
  expect_equal(un_xlex("SOME_FUNCTION():SOME.NAME"),
               tribble(~level,       ~type,          ~token,
                           0L,  "function", "SOME_FUNCTION",
                           0L,  "fun_open",             "(",
                           0L, "fun_close",             ")",
                           0L,  "operator",             ":",
                           0L,      "name",     "SOME.NAME"))
})

test_that("sheet names are recognised by the terminal exclamation mark '!'", {
  expect_equal(un_xlex("Sheet1!A1"),
               tribble(~level,   ~type,    ~token,
                           0L, "sheet", "Sheet1!",
                           0L,   "ref",      "A1"
))
  expect_equal(un_xlex("'Sheet 1'!A1"),
               tribble(~level,   ~type,       ~token,
                           0L, "sheet", "'Sheet 1'!",
                           0L,   "ref",         "A1"))
  expect_equal(un_xlex("'It''s a sheet'!A1"),
               tribble(~level,   ~type,             ~token,
                           0L, "sheet", "'It''s a sheet'!",
                           0L,   "ref",               "A1"))
})

test_that("both sheets in 3d formulas are returned in a single 'sheet' token", {
  expect_equal(un_xlex("Sheet1:Sheet2!A1"),
               tribble(~level,   ~type,           ~token,
                           0L, "sheet", "Sheet1:Sheet2!",
                           0L,   "ref",             "A1"))
  expect_equal(un_xlex("'Sheet 1:Sheet 2'!A1"),
               tribble(~level,   ~type,               ~token,
                           0L, "sheet", "'Sheet 1:Sheet 2'!",
                           0L,   "ref",                 "A1"))
})

test_that("file paths/indexes are included in the 'sheet' token", {
  expect_equal(un_xlex("[1]Sheet1!A1"),
               tribble(~level,   ~type,       ~token,
                           0L, "sheet", "[1]Sheet1!",
                           0L,   "ref",         "A1"))
  expect_equal(un_xlex("'[1]Sheet 1'!A1"),
               tribble(~level,   ~type,          ~token,
                           0L, "sheet", "'[1]Sheet 1'!",
                           0L,   "ref",            "A1"))
  expect_equal(un_xlex("'C:\\My Documents\\[file.xlsx]Sheet1'!A1"),
               tribble(~level,   ~type,                                   ~token,
                           0L, "sheet", "'C:\\My Documents\\[file.xlsx]Sheet1'!",
                           0L,   "ref",                                     "A1"))
})

test_that("functions are recognised by a terminal open-parenthesis '('", {
  expect_equal(un_xlex("MAX(1,2)"),
               tribble(~level,       ~type, ~token,
                           0L,  "function",  "MAX",
                           0L,  "fun_open",    "(",
                           1L,    "number",    "1",
                           1L, "separator",    ",",
                           1L,    "number",    "2",
                           0L, "fun_close",    ")"))
  expect_equal(un_xlex("_xll.MY_CUSTOM_FUNCTION()"),
               tribble(~level,       ~type,                    ~token,
                           0L,  "function", "_xll.MY_CUSTOM_FUNCTION",
                           0L,  "fun_open",                       "(",
                           0L, "fun_close",                       ")"))
})

test_that("named ranges/formulas are returned", {
  expect_equal(un_xlex("MY_NAMED_RANGE"),
               tribble(~level,  ~type,           ~token,
                           0L, "name", "MY_NAMED_RANGE"))
})

test_that("lookalike ref/function/name are typed correctly", {
  expect_equal(xlex("XFD1")$type[1],     "ref")
  expect_equal(xlex("XFE1")$type[1],     "name")
  expect_equal(xlex("A1048576")$type[1], "ref")
  expect_equal(xlex("A1048577")$type[1], "name")
  expect_equal(xlex("LOG10")$type[1],    "ref")
  expect_equal(xlex("LOG10()")$type[1],  "function")
  expect_equal(xlex("LOG:LOG")$type[1],  "ref")
  expect_equal(xlex("LOG")$type[1],      "name")
  expect_equal(xlex("LOG()")$type[1],    "function")
  expect_equal(xlex("A1.2!A1")$type[1],  "sheet")
})

test_that("text is returned in a single token including escaped quotes", {
  expect_equal(un_xlex("\"Some text\""),
               tribble(~level,  ~type,          ~token,
                           0L, "text", "\"Some text\""))
  expect_equal(un_xlex("\"Some \"\"text\"\"\""),
               tribble(~level,  ~type,                  ~token,
                           0L, "text", "\"Some \"\"text\"\"\""))
})

# Numbers are signed where it makes sense, and can be scientific
test_that("numbers are signed where it makes sense, and can be scientific", {
  expect_equal(un_xlex("1"),
               tribble(~level,    ~type, ~token,
                           0L, "number",    "1"))
  expect_equal(un_xlex("1.2"),
               tribble(~level,    ~type, ~token,
                           0L, "number",  "1.2"))
  expect_equal(un_xlex("-1"),
               tribble(~level,    ~type, ~token,
                           0L, "number",   "-1"))
  expect_equal(un_xlex("-1-1"),
               tribble(~level,      ~type, ~token,
                           0L,   "number",   "-1",
                           0L, "operator",    "-",
                           0L,   "number",    "1"))
  expect_equal(un_xlex("-1+-1"),
               tribble(~level,      ~type, ~token,
                           0L,   "number",   "-1",
                           0L, "operator",    "+",
                           0L,   "number",  "-1" ))
  expect_equal(un_xlex("MAX(-1-1)"),
               tribble(~level,       ~type, ~token,
                           0L,  "function",  "MAX",
                           0L,  "fun_open",    "(",
                           1L,  "operator",    "-",
                           1L,    "number",    "1",
                           1L,  "operator",    "-",
                           1L,    "number",    "1",
                           0L, "fun_close",   ")" ))
  expect_equal(un_xlex("-1.2E-3"),
               tribble(~level,    ~type,    ~token,
                           0L, "number", "-1.2E-3"))
})

test_that("lookalike bool/function/name are typed correctly", {
  expect_equal(xlex("TRUE")$type[1],    "bool")
  expect_equal(xlex("TRUEISH")$type[1], "name")
  expect_equal(xlex("TRUE!A1")$type[1], "sheet")
  expect_equal(xlex("TRUE()")$type[1],  "function")
})

test_that("Errors are tagged 'error'", {
  expect_equal(xlex("#DIV/0!")$type[1], "error")
  expect_equal(xlex("#N/A")$type[1],    "error")
  expect_equal(xlex("#NAME?")$type[1],  "error")
  expect_equal(xlex("#NULL!")$type[1],  "error")
  expect_equal(xlex("#NUM!")$type[1],   "error")
  expect_equal(xlex("#REF!")$type[1],   "error")
  expect_equal(xlex("#VALUE!")$type[1], "error")
})

test_that("multi-char operators are treated as single tokens", {
  expect_equal(xlex("1<>2")$token[2], "<>")
  expect_equal(xlex("1<=2")$token[2], "<=")
  expect_equal(xlex("1<2")$token[2],  "<")
  expect_equal(xlex("1=2")$token[2],  "=")
  expect_equal(xlex("1&2")$token[2],  "&")
  expect_equal(xlex("1 2")$token[2],  " ")
  expect_equal(xlex("1%")$token[2],   "%")
})

test_that("commas are correctly tagged as operators (union) or separators", {
  expect_equal(un_xlex("A1,B2"),
               # invalid formula, defaults to 'union' to avoid a crash
               tribble(~level,      ~type, ~token,
                           0L,      "ref",   "A1",
                           0L, "operator",    ",",
                           0L,      "ref",   "B2"))
  expect_equal(un_xlex("(A1,B2)"),
               tribble(~level,         ~type, ~token,
                           0L,  "paren_open",    "(",
                           1L,         "ref",   "A1",
                           1L,    "operator",    ",",
                           1L,         "ref",   "B2",
                           0L, "paren_close",    ")"))
  expect_equal(un_xlex("MAX(A1,B2)"),
               tribble(~level,       ~type, ~token,
                           0L,  "function",  "MAX",
                           0L,  "fun_open",    "(",
                           1L,       "ref",   "A1",
                           1L, "separator",    ",",
                           1L,       "ref",   "B2",
                           0L, "fun_close",    ")"))
  expect_equal(un_xlex("SMALL((A1,B2),1)"),
               tribble(~level,         ~type,  ~token,
                           0L,    "function", "SMALL",
                           0L,    "fun_open",     "(",
                           1L,  "paren_open",     "(",
                           2L,         "ref",    "A1",
                           2L,    "operator",     ",",
                           2L,         "ref",    "B2",
                           1L, "paren_close",     ")",
                           1L,   "separator",     ",",
                           1L,      "number",     "1",
                           0L,   "fun_close",     ")"))
  expect_equal(un_xlex("{1,2;3,4}"),
               tribble(~level,         ~type, ~token,
                           0L,  "open_array",    "{",
                           1L,      "number",    "1",
                           1L,   "separator",    ",",
                           1L,      "number",    "2",
                           1L,   "separator",    ";",
                           1L,      "number",    "3",
                           1L,   "separator",    ",",
                           1L,      "number",    "4",
                           0L, "close_array",    "}"))
})

test_that("the level increases within but not at parentheses and arrays", {
  expect_equal(un_xlex("MAX(MIN(1,2),3)"),
               tribble(~level,       ~type, ~token,
                           0L,  "function",  "MAX",
                           0L,  "fun_open",    "(",
                           1L,  "function",  "MIN",
                           1L,  "fun_open",    "(",
                           2L,    "number",    "1",
                           2L, "separator",    ",",
                           2L,    "number",    "2",
                           1L, "fun_close",    ")",
                           1L, "separator",    ",",
                           1L,    "number",    "3",
                           0L, "fun_close",  ")"))
  expect_equal(un_xlex("{1,2;3,4}"),
               tribble(~level,         ~type, ~token,
                           0L,  "open_array",    "{",
                           1L,      "number",    "1",
                           1L,   "separator",    ",",
                           1L,      "number",    "2",
                           1L,   "separator",    ";",
                           1L,      "number",    "3",
                           1L,   "separator",    ",",
                           1L,      "number",    "4",
                           0L, "close_array",    "}"))
  expect_equal(un_xlex("1*(2+3)"),
               tribble(~level,         ~type, ~token,
                           0L,      "number",    "1",
                           0L,    "operator",    "*",
                           0L,  "paren_open",    "(",
                           1L,      "number",    "2",
                           1L,    "operator",    "+",
                           1L,      "number",    "3",
                           0L, "paren_close",    ")"))
})

test_that("whole structured references are returned in a single token", {
  expect_equal(un_xlex("[@col2]"),
                tribble(~level,            ~type,    ~token,
                            0L, "structured_ref", "[@col2]"))
  expect_equal(un_xlex("SUM([col22])"),
               tribble(~level,            ~type,    ~token,
                           0L,       "function",     "SUM",
                           0L,       "fun_open",       "(",
                           1L, "structured_ref", "[col22]",
                           0L,      "fun_close",       ")"))
  expect_equal(un_xlex("Table1[col1]"),
               tribble(~level,            ~type,         ~token,
                           0L, "structured_ref", "Table1[col1]" ))
  expect_equal(un_xlex("Table1[[col1]:[col2]]"),
               tribble(~level,            ~type,                  ~token,
                           0L, "structured_ref", "Table1[[col1]:[col2]]"))
  expect_equal(un_xlex("Table1[#Headers]"),
               tribble(~level,            ~type,             ~token,
                           0L, "structured_ref", "Table1[#Headers]"))
  expect_equal(un_xlex("Table1[[#Headers],[col1]]"),
               tribble(~level,            ~type,                       ~token,
                           0L, "structured_ref", "Table1[[#Headers],[col1]]"))
  expect_equal(un_xlex("Table1[[#Headers],[col1]:[col2]]"),
               tribble(~level,            ~type,                              ~token,
                           0L, "structured_ref", "Table1[[#Headers],[col1]:[col2]]"))
})

test_that("whole dynamic data exchange calls are returned in a single token", {
  expect_equal(un_xlex("[1]!'DDE_parameter=1'"),
                tribble(~level, ~type,                  ~token,
                            0L, "DDE", "[1]!'DDE_parameter=1'"))
  expect_equal(un_xlex("'Quote'|'NYSE'!ZAXX"),
               tribble(~level, ~type,                ~token,
                           0L, "DDE", "'Quote'|'NYSE'!ZAXX"))
})

test_that("argument type and length is enforced", {
  expect_error(xlex(c("A1", "B2")), "'x' must be a character vector of length 1")
  expect_error(xlex(1), "'x' must be a character vector of length 1")
})

test_that("unrecognised tokens are returned as 'other'", {
  expect_equal(xlex(".")$type, "other")
})

test_that("print.xlex() doesn't error", {
  expect_error(print(xlex("1")), NA)
  expect_error(print(xlex("1"), pretty = FALSE), NA)
})
