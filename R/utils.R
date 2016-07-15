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

standardise_sheet <- function(sheets, sheet_names) {
  if (is.numeric(sheets)) {
    if (max(sheets) > length(sheet_names)) {
      stop("Only ", length(sheet_names), " sheet(s) found.", call. = FALSE)
    }
    floor(sheets) - 1L
  } else if (is.character(sheets)) {
    indices <- match(sheets, sheet_names)
    if (anyNA(indices)) {
      stop("Sheet(s) not found: ", paste(sheets[is.na(indices)], collapse = "\", \""),
           call. = FALSE)
    }
    indices - 1L
  } else {
    stop("`sheet` must be either an integer or a string.", call. = FALSE)
  }
}

