#' @importFrom magrittr %>%
#' @export
magrittr::`%>%`

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

xlsx_sheets <- function(path) {
  relationships <- 
    unz(path, "xl/_rels/workbook.xml.rels") %>%
    xml2::read_xml() %>%
    xml2::xml_ns_strip() %>%
    xml2::xml_find_all("Relationship") %>%
    xml2::xml_attrs() %>%
    purrr::transpose() %>%
    purrr::map(unlist) %>% 
    tibble::as_tibble()
  sheets <-
    unz(path, "xl/workbook.xml") %>%
    xml2::read_xml() %>%
    xml2::xml_ns_strip() %>%
    xml2::xml_find_all("sheets/sheet") %>%
    xml2::xml_attrs() %>%
    purrr::map(as.list) %>%
    purrr::map(tibble::as_tibble) %>%
    dplyr::bind_rows()
  comments_path <- function(rels) {
    if (zip_has_file(path, rels)) {
      nodes <- 
        unz(path, rels) %>%
        xml2::read_xml() %>% 
        xml2::xml_ns_strip() %>%
        xml2::xml_find_all("Relationship") %>%
        .[xml2::xml_has_attr(., "Type")] %>%
        .[xml2::xml_has_attr(., "Target")]
      Type <- xml2::xml_attr(nodes, "Type")
      Target <- xml2::xml_attr(nodes, "Target")
      Target <- Target[grepl("/comments", Type)]
      if (length(Target) > 0) {
        substr(Target, 1, 2) <- "xl"
        return(Target)
      }
    }
    return(as.character(NA))
  }
  dplyr::inner_join(relationships, sheets, by = c("Id" = "id")) %>% # Don't use sheetId -- out of sync with rId
    dplyr::filter(grepl("worksheets/", Target, fixed = TRUE)) %>% # Ignore chartsheets etc.
    dplyr::mutate(id = as.integer(gsub("rId", "", Id)),
                  rels = paste0("xl/worksheets/_rels/sheet", id, ".xml.rels"),
                  comments_path = vapply(rels, comments_path, character(1))) %>%
    dplyr::arrange(id) %>%
    dplyr::mutate(index = row_number()) %>%
    dplyr::select(index, name, comments_path)
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

