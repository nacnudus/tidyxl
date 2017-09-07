context("xlsx_validation()")

library(tibble)

test_that("xlsx_validation() doesn't fail when there aren't any rules", {
  expect_error(xlsx_validation("./default.xlsx"), NA)
})

test_that("Data validation rules are returned correctly", {
  expect_equal(xlsx_validation("./examples.xlsx")$Sheet1,
                 tribble(~ref,        ~type,            ~operator,             ~formula1,             ~formula2,  ~allow_blank, ~show_input_message,   ~prompt_title,   ~prompt_body, ~show_error_message,  ~error_title,   ~error_body, ~error_symbol,
                       "A106",      "whole",            "between",                   "0",                   "9",          TRUE,                TRUE, "message title", "message body",                TRUE, "error title",  "error body",        "stop",
                       "A107",    "decimal",         "notBetween",                   "0",                   "9",         FALSE,               FALSE,   NA_character_,  NA_character_,               FALSE, NA_character_, NA_character_,        "stop",
                       "A108",       "list",        NA_character_,              "$B$108",         NA_character_,          TRUE,                TRUE,   NA_character_,  NA_character_,                TRUE, NA_character_, NA_character_,     "warning",
                       "A109",       "list",        NA_character_,              "$B$108",         NA_character_,          TRUE,                TRUE,   NA_character_,  NA_character_,                TRUE, NA_character_, NA_character_, "information",
                       "A110",       "date",            "between", "2017-01-01 00:00:00", "2017-01-09 09:00:00",          TRUE,                TRUE,   NA_character_,  NA_character_,                TRUE, NA_character_, NA_character_,        "stop",
                       "A111",       "time",            "between",            "00:00:00",            "09:00:00",          TRUE,                TRUE,   NA_character_,  NA_character_,                TRUE, NA_character_, NA_character_,        "stop",
                       "A112", "textLength",            "between",                   "0",                   "9",          TRUE,                TRUE,   NA_character_,  NA_character_,                TRUE, NA_character_, NA_character_,        "stop",
                       "A113",     "custom",        NA_character_,     "A113<=LEN(B113)",         NA_character_,          TRUE,                TRUE,   NA_character_,  NA_character_,                TRUE, NA_character_, NA_character_,        "stop",
                       "A114",      "whole",         "notBetween",                   "0",                   "9",          TRUE,                TRUE,   NA_character_,  NA_character_,                TRUE, NA_character_, NA_character_,        "stop",
             "A115,A121:A122",      "whole",              "equal",                   "0",         NA_character_,          TRUE,                TRUE,   NA_character_,  NA_character_,                TRUE, NA_character_, NA_character_,        "stop",
                       "A116",      "whole",           "notEqual",                   "0",         NA_character_,          TRUE,                TRUE,   NA_character_,  NA_character_,                TRUE, NA_character_, NA_character_,        "stop",
                       "A117",      "whole",        "greaterThan",                   "0",         NA_character_,          TRUE,                TRUE,   NA_character_,  NA_character_,                TRUE, NA_character_, NA_character_,        "stop",
                       "A118",      "whole",           "lessThan",                   "0",         NA_character_,          TRUE,                TRUE,   NA_character_,  NA_character_,                TRUE, NA_character_, NA_character_,        "stop",
                       "A119",      "whole", "greaterThanOrEqual",                   "0",         NA_character_,          TRUE,                TRUE,   NA_character_,  NA_character_,                TRUE, NA_character_, NA_character_,        "stop",
                       "A120",      "whole",    "lessThanOrEqual",                   "0",         NA_character_,          TRUE,                TRUE,   NA_character_,  NA_character_,                TRUE, NA_character_, NA_character_,        "stop"))
})
