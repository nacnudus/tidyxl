globalVariables(c(".",
                  "numFmtId",
                  "Target",
                  "Id",
                  "id",
                  "rels",
                  "row_number",
                  "index",
                  "name",
                  "hasArg"))

check_file <- function(path, check_filetype = TRUE) {
  if (!file.exists(path)) {
    stop("'", path, "' does not exist",
      if (!is_absolute_path(path))
        paste0(" in current working directory ('", getwd(), "')"),
      ".",
      call. = FALSE)
  }

  if (check_filetype && !maybe_xlsx(path)) {
    stop("The file format does not appear to be xlsx, xlsm, xltx or xltm,",
         "\neven if the filename extension says so.",
         "\n  ", path,
         "\n", "You could try converting it with a spreadsheet application.",
         "\n", "If you need tidyxl to support files of this type, then please ",
         "\n", "open an issue at https://github.com/nacnudus/tidyxl/issues.",
         call. = FALSE)
  }

  normalizePath(path, "/", mustWork = FALSE)
}

is_absolute_path <- function(path) {
  grepl("^(/|[A-Za-z]:|\\\\|~)", path)
}

standardise_sheet <- function(sheets, all_sheets) {
  if (is.numeric(sheets)) {
    if (max(sheets) > max(all_sheets$order)) {
      stop("Only ", max(all_sheets$order), " sheet(s) found.", call. = FALSE)
    }
    if (any(!(sheets %in% all_sheets$order))) {
      warning("Only worksheets (not chartsheets) were imported", call. = FALSE)
    }
    all_sheets[all_sheets$order %in% sheets, ]
  } else if (is.character(sheets)) {
    indices <- match(sheets, all_sheets$name)
    if (anyNA(indices)) {
      stop("Sheets not found: \"",
           paste(sheets[is.na(indices)],
                 collapse = "\", \""),
           "\"",
           call. = FALSE)
    }
    all_sheets[indices, ]
  } else {
    stop("Argument `sheet` must be either an integer or a string.",
         call. = FALSE)
  }
}

check_sheets <- function(sheets, path) {
  all_sheets <- utils_xlsx_sheet_files(path)
  if (anyNA(sheets)) {
    if (length(sheets) > 1) {
      warning("Argument 'sheets' included NAs, which were discarded.")
      sheets <- sheets[!is.na(sheets)]
      if (length(sheets) == 0) {
        stop("All elements of argument 'sheets' were discarded.")
      }
    } else {
      sheets <- all_sheets$order
    }
  }
  standardise_sheet(sheets, all_sheets)
}

utils_xlsx_sheet_files <- function(path) {
  out <- xlsx_sheet_files_(path)
  # Standardise /xl/worksheets/sheet1.xml and worksheets/sheet1.xml
  out$sheet_path <- gsub("^/", "", out$sheet_path)
  out$sheet_path <- gsub("^xl/", "", out$sheet_path)
  out$sheet_path <- paste0("xl/", out$sheet_path)
  out$order <- order(out$rId)
  out <- out[out$order, ]
  # Omit chartsheets
  out <- out[grepl("^xl/worksheets\\.*", out$sheet_path), ]
  out
}
