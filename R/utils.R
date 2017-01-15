globalVariables(c(".",
                  "numFmtId",
                  "Target",
                  "Id",
                  "id",
                  "rels",
                  "row_number",
                  "index",
                  "name"))

check_file <- function(path) {
  if (!file.exists(path)) {
    stop("'", path, "' does not exist",
      if (!is_absolute_path(path))
        paste0(" in current working directory ('", getwd(), "')"),
      ".",
      call. = FALSE)
  }

  excel_format(path)

  normalizePath(path, "/", mustWork = FALSE)
}

is_absolute_path <- function(path) {
  grepl("^(/|[A-Za-z]:|\\\\|~)", path)
}

excel_format <- function(path) {
  ext <- tolower(tools::file_ext(path))

  switch(ext,
    xls = stop("Not implemented for xls files.", call. = FALSE),
    xlsx = "xlsx",
    xlsm = "xlsx",
    stop("Unknown format .", ext, call. = FALSE)
  )
}

standardise_sheet <- function(sheets, all_sheets) {
  if (is.numeric(sheets)) {
    if (max(sheets) > nrow(all_sheets)) {
      stop("Only ", nrow(all_sheets), " sheet(s) found.", call. = FALSE)
    }
    all_sheets[sheets, ]
  } else if (is.character(sheets)) {
    indices <- match(sheets, all_sheets$name)
    if (anyNA(indices)) {
      stop("Sheet(s) not found: ", paste(sheets[is.na(indices)], collapse = "\", \""),
           call. = FALSE)
    }
    all_sheets[indices, ]
  } else {
    stop("Argument `sheet` must be either an integer or a string.",
         call. = FALSE)
  }
}

xlsx_sheets <- function(path) {
  out <- xlsx_sheets_(path)
  out[order(out$index), ]
}
