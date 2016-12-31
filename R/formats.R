#' @title Import xlsx (Excel) formatting into a list structure for lookups.
#'
#' @description
#' \code{formats} imports formatting information from xlsx (Excel) spreadsheets
#' into a list structure.  The top level of the list is divided into two parts:
#' \code{$style}, which contains formats applied in a general way by cell
#' 'styles', and \code{$local}, which contains formats applied to individual
#' cells.  This structure reflects how a cell may be formatted according to a
#' general 'style', yet also differ from the 'style' by some local modification,
#' e.g. by making it bold.
#'
#' @param path Path to the xlsx file.
#'
#' @details
#' \code{\link{contents}} Returns, among other things, keys for each cell into
#' the \code{$style} and \code{$local} substructures of \code{formats}' output.
#' These keys can be used to retrieve the style of each cell by subsetting the
#' relevant part of \code{formats}' output.  See how this is done in the
#' examples.
#'
#' Colours may be recorded in any of three ways: a hexadecimal RGB string with
#' or without alpha, an 'indexed' colour, and an index into a 'theme'.
#' \code{formats} dereferences 'indexed' and 'theme' colours to their hexadecimal
#' RGB string representation, and standardises all RGB strings to have an alpha
#' channel in the first two characters.  The 'index' and the 'theme' name are
#' still provided.  To filter by an RGB string, you might to look up the RGB
#' values in a spreadsheet program (e.g. Excel, LibreOffice, Gnumeric), and use
#' the \code{\link{rgb}} function to convert these to a hexadecimal string.
#'
#' \code{formats} is currently quite slow for spreadsheets with complex
#' formatting.  It is planned to implement \code{formats} in C++ to speed it up.
#'
#' @export
#' @examples
#' examples <- system.file("extdata/examples.xlsx", package = "tidyxl")
#' examples_formats <- formats(examples)
#' str(examples_formats)
#'
#' # The formats of particular cells can be retrieved like this:
#' sheet1 <- contents(examples, "Sheet1")[[1]]
#' examples_formats$style$font$bold[sheet1$style_format_id]
#' examples_formats$local$font$bold[sheet1$local_format_id]
#'
#' # To filter for cells of a particular format, first filter the formats to get
#' # the relevant indices, and then filter the cells by those indices.
#' bold_indices <- which(examples_formats$local$font$bold)
#' sheet1[sheet1$local_format_id %in% bold_indices, ]
formats <- function(path) {
  styleSheet <-
    path %>%
    unz("xl/styles.xml") %>%
    xml2::read_xml() %>%
    xml2::xml_ns_strip()
  if (zip_has_file(path, "xl/theme/theme1.xml")) {
    theme <-
      path %>%
      unz("xl/theme/theme1.xml") %>%
      xml2::read_xml()
    colour_scheme <- clrScheme(theme)
  } else {
    colour_scheme <- data.frame(name = character(12),
                                rgb = character(12),
                                stringsAsFactors = FALSE)
  }
  numFmts <- get_numFmts(styleSheet)
  fonts <- get_fonts(styleSheet)
  fills <- get_fills(styleSheet)
  borders <- get_borders(styleSheet)
  normal <-
    styleSheet %>%
    xml2::xml_find_first("cellStyleXfs") %>%
    xml2::xml_child(search = 1) %>%
    xf(numFmts, fonts, fills, borders)
  cellStyleXfs <-
    styleSheet %>%
    xml2::xml_find_first("cellStyleXfs") %>%
    xml2::xml_children() %>%
    purrr::map(cellStyleXf,
               normal = normal,
               numFmts = numFmts,
               fonts = fonts,
               fills = fills,
               borders = borders)
  cellXfs <-
    styleSheet %>%
    xml2::xml_find_first("cellXfs") %>%
    xml2::xml_children() %>%
    purrr::map(cellXf,
               cellStyleXfs = cellStyleXfs,
               numFmts = numFmts,
               fonts = fonts,
               fills = fills,
               borders = borders)
  list(style = cellStyleXfs, local = cellXfs) %>%
    purrr::map(transpose_formats, colour_scheme = colour_scheme)
}

get_numFmts <- function(styleSheet) {
  numFmt_defaults <-
      tibble::frame_data(
        ~numFmtId, ~formatCode,
        0, "General",
        1, "0",
        2, "0.00",
        3, "#,##0",
        4, "#,##0.00",
        9, "0%",
        10, "0.00%",
        11, "0.00E+00",
        12, "# ?/?",
        13, "# ??/??",
        14, "mm-dd-yy",
        15, "d-mmm-yy",
        16, "d-mmm",
        17, "mmm-yy",
        18, "h:mm AM/PM",
        19, "h:mm:ss AM/PM",
        20, "h:mm",
        21, "h:mm:ss",
        22, "m/d/yy h:mm",
        37, "#,##0 ;(#,##0)",
        38, "#,##0 ;[Red](#,##0)",
        39, "#,##0.00;(#,##0.00)",
        40, "#,##0.00;[Red](#,##0.00)",
        45, "mm:ss",
        46, "[h]:mm:ss",
        47, "mmss.0",
        48, "##0.0E+0",
        49, "@"
      ) %>%
      dplyr::mutate(numFmtId = as.integer(numFmtId))
  parse_numFmt <- function(numFmt) {
    list(
      numFmtId = as.integer(xml2::xml_attr(numFmt, "numFmtId")),
      formatCode = xml2::xml_attr(numFmt, "formatCode")
    )
  }
  numFmtsNode <- xml2::xml_find_first(styleSheet, "numFmts")
  if (class(numFmtsNode) == "xml_missing") {
    numFmt_custom <- NULL
  } else {
    numFmt_custom <-
      xml2::xml_find_first(styleSheet, "numFmts") %>%
      xml2::xml_children() %>%
      purrr::map(parse_numFmt) %>%
      dplyr::bind_rows()
  }
  dplyr::bind_rows(numFmt_defaults, numFmt_custom) %>%
    dplyr::right_join(data.frame(numFmtId = seq(0, max(.$numFmtId))),
                      by = "numFmtId") %>%
    .$formatCode
}

# ECMA Part I p.1764, with the first 00 replaced by FF for consistency with
# what argb colours seem to do, (i.e. using FF to mean opaque, not 00).
indexed <- c("FF000000",
             "FFFFFFFF",
             "FFFF0000",
             "FF00FF00",
             "FF0000FF",
             "FFFFFF00",
             "FFFF00FF",
             "FF00FFFF",
             "FF000000",
             "FFFFFFFF",
             "FFFF0000",
             "FF00FF00",
             "FF0000FF",
             "FFFFFF00",
             "FFFF00FF",
             "FF00FFFF",
             "FF800000",
             "FF008000",
             "FF000080",
             "FF808000",
             "FF800080",
             "FF008080",
             "FFC0C0C0",
             "FF808080",
             "FF9999FF",
             "FF993366",
             "FFFFFFCC",
             "FFCCFFFF",
             "FF660066",
             "FFFF8080",
             "FF0066CC",
             "FFCCCCFF",
             "FF000080",
             "FFFF00FF",
             "FFFFFF00",
             "FF00FFFF",
             "FF800080",
             "FF800000",
             "FF008080",
             "FF0000FF",
             "FF00CCFF",
             "FFCCFFFF",
             "FFCCFFCC",
             "FFFFFF99",
             "FF99CCFF",
             "FFFF99CC",
             "FFCC99FF",
             "FFFFCC99",
             "FF3366FF",
             "FF33CCCC",
             "FF99CC00",
             "FFFFCC00",
             "FFFF9900",
             "FFFF6600",
             "FF666699",
             "FF969696",
             "FF003366",
             "FF339966",
             "FF003300",
             "FF333300",
             "FF993300",
             "FF993366",
             "FF333399",
             "FF333333",
             "FFFFFFFF", # System foreground
             "FF000000", # System background
             (NA),
             NA,
             NA,
             NA,
             NA,
             NA,
             NA,
             NA,
             NA,
             NA,
             NA,
             NA,
             NA,
             NA,
             NA,
             "FF000000") # Undocumented comment text in black

clrScheme <- function(theme) {
  elements <-
    theme %>%
    xml2::xml_find_first("a:themeElements") %>%
    xml2::xml_find_first("a:clrScheme") %>%
    xml2::xml_children()
  name <- elements %>% xml2::xml_name()
  sysClr  <- xml2::xml_find_first(elements, "a:sysClr") %>% xml2::xml_attr("lastClr")
  srgbClr   <- xml2::xml_find_first(elements, "a:srgbClr") %>% xml2::xml_attr("val")
  # Take either sysClr or srgbClr, whichever is present
  sysClr_NA <- which(is.na(sysClr))
  rgb <- sysClr
  rgb[sysClr_NA] <- srgbClr[sysClr_NA]
  rgb <- paste0("FF", rgb) # Add opaque alpha for consistency with srgb
  # Flip the order of the first two pairs
  out <- data.frame(name = name, rgb = rgb,
                    stringsAsFactors = FALSE)[c(2, 1, 4, 3, 5:12), ]
  out
}

decode_color <- function(color, colour_scheme) {
  # Decodes 'indexed' and 'theme' columns, combining the results with the
  # existing 'rgb' column.
  color$rgb <- ifelse(is.na(color$rgb), indexed[color$indexed], color$rgb)
  color$rgb <- ifelse(is.na(color$rgb), colour_scheme$rgb[color$theme], color$rgb)
  color$rgb <- ifelse(is.na(color$rgb), "FFFFFFFF", color$rgb) # 'automatic' font
  color$theme <- colour_scheme$name[color$theme]
  color
}

get_fonts <- function(styleSheet) {
  font_color <- function(font) {
    color <- xml2::xml_find_first(font, "color")
    if (class(color) == "xml_missing") {
      list(rgb = as.character(NA),
           indexed = as.integer(NA),
           theme = as.character(NA))
    } else {
      list(rgb = xml2::xml_attr(color, "rgb"),
           indexed = as.integer(xml2::xml_attr(color, "indexed")) + 1,
           theme = as.integer(xml2::xml_attr(color, "theme")) + 1)
    }
  }
  font <- function(node) {
    list(bold = class(xml2::xml_find_first(node, "b")) != "xml_missing",
         italic = class(xml2::xml_find_first(node, "i")) != "xml_missing",
         underline = class(xml2::xml_find_first(node, "u")) != "xml_missing",
         strike = class(xml2::xml_find_first(node, "strike")) != "xml_missing",
         size = xml2::xml_find_first(node, "sz") %>% xml2::xml_attr("val") %>% as.integer,
         color = font_color(node),
         name = xml2::xml_find_first(node, "name") %>% xml2::xml_attr("val"),
         family = xml2::xml_find_first(node, "family") %>% xml2::xml_attr("val") %>% as.integer,
         scheme = xml2::xml_find_first(node, "scheme") %>% xml2::xml_attr("val"))
  }
  fontsNode <- xml2::xml_find_first(styleSheet, "fonts")
  if (class(fontsNode) == "xml_missing") {
    return(NULL)
  } else {
    styleSheet %>%
    xml2::xml_find_first("fonts") %>%
    xml2::xml_children() %>%
    purrr::map(font) %>%
    return
  }
}

get_fills <- function(styleSheet) {
  fill_color <- function(patternFill, layer = c("bgColor", "fgColor")) {
    xgColor <- xml2::xml_find_first(patternFill, layer)
    if (class(xgColor) == "xml_missing") {
      list(rgb = as.character(NA),
           indexed = as.integer(NA),
           theme = as.character(NA))
    } else {
      list(rgb = xml2::xml_attr(xgColor, "rgb"),
           indexed = as.integer(xml2::xml_attr(xgColor, "indexed")) + 1,
           theme = as.integer(xml2::xml_attr(xgColor, "theme")) + 1)
    }
  }
  fill <- function(node) {
    patternFill <- xml2::xml_find_first(node, "patternFill")
    if (class(patternFill) == "xml_missing") {
      patternType <- NA
      fgColor <- bgColor <- list(rgb = as.character(NA),
                                 indexed = as.integer(NA),
                                 theme = as.character(NA))
    } else {
      patternType <- xml2::xml_attr(patternFill, "patternType")
      bgColor <- fill_color(patternFill, "bgColor")
      fgColor <- fill_color(patternFill, "fgColor")
    }
    list(patternType = patternType,
         bgColor = bgColor,
         fgColor = fgColor)
  }
  fillsNode <- xml2::xml_find_first(styleSheet, "fills")
  if (class(fillsNode) == "xml_missing") {
    return(NULL)
  } else {
    styleSheet %>%
    xml2::xml_find_first("fills") %>%
    xml2::xml_children() %>%
    purrr::map(fill) %>%
    return
  }
}

get_borders <- function(styleSheet) {
  border_position <- function(border,
                              position = c("left", "right", "top",
                                           "bottom", "diagonal")) {
    node <- xml2::xml_find_first(border, position)
    if (class(node) == "xml_missing") {
      list(style = character(),
           color = list(rgb = character(),
                        indexed = integer(),
                        theme = integer())) %>%
      return
    } else {
      color <- xml2::xml_find_first(node, "color")
      if (is.na(color)) {
        list(style = character(),
             color = list(rgb = character(),
                          indexed = integer(),
                          theme = integer())) %>%
        return
      } else {
        list(style = xml2::xml_attr(node, "style"),
             color = list(rgb = xml2::xml_attr(color, "rgb"),
                          indexed = as.integer(xml2::xml_attr(color, "indexed")) + 1,
                          theme = as.integer(xml2::xml_attr(color, "theme")) + 1)) %>%
        return
      }
    }
  }
  border <- function(node) {
    list(left = border_position(node, "left"),
         right = border_position(node, "right"),
         top = border_position(node, "top"),
         bottom = border_position(node, "bottom"),
         diagonal = border_position(node, "diagonal"))
  }
  bordersNode <- xml2::xml_find_first(styleSheet, "borders")
  if (class(bordersNode) == "xml_missing") {
    return(NULL)
  } else {
    styleSheet %>%
    xml2::xml_find_first("borders") %>%
    xml2::xml_children() %>%
    purrr::map(border) %>%
    return
  }
}

xf <- function(node, numFmts, fonts, fills, borders) {
  alignment <- xml2::xml_find_first(node, "alignment")
  protection <- xml2::xml_find_first(node, "protection")
  list(number_format = numFmts[as.integer(xml2::xml_attr(node, "numFmtId")) + 1],
       font = fonts[[as.integer(xml2::xml_attr(node, "fontId")) + 1]],
       fill = fills[[as.integer(xml2::xml_attr(node, "fillId")) + 1]],
       border = borders[[as.integer(xml2::xml_attr(node, "borderId")) + 1]],
       alignment = list(horizontal = xml2::xml_attr(alignment, "horizontal"),
                        vertical = xml2::xml_attr(alignment, "vertical"),
                        indent = xml2::xml_attr(alignment, "indent"),
                        wrap_text = xml2::xml_attr(alignment, "wrapText"),
                        justify_last_line = xml2::xml_attr(alignment, "justifyLastLine"),
                        shrink_to_fit = xml2::xml_attr(alignment, "shrinkToFit"),
                        reading_order = xml2::xml_attr(alignment, "readingOrder")
                        ),
       protection = list(locked = xml2::xml_attr(protection, "locked"),
                         hidden = xml2::xml_attr(protection, "hidden")))
}

cellStyleXf <- function(node, normal, numFmts, fonts, fills, borders) {
  alignment <- xml2::xml_find_first(node, "alignment")
  protection <- xml2::xml_find_first(node, "protection")
  out <- normal # Start from normal style, then modify
  bool_cellStyle <- function(attribute) { # Whether to modify something
    bool <- as.integer(attribute)
    bool | is.na(bool)
  }
  if (bool_cellStyle(xml2::xml_attr(node, "applyNumberFormat"))) {
    out$number_format <- numFmts[as.integer(xml2::xml_attr(node, "numFmtId")) + 1]
  }
  if (bool_cellStyle(xml2::xml_attr(node, "applyFont"))) {
    out$font <- fonts[[as.integer(xml2::xml_attr(node, "fontId")) + 1]]
  }
  if (bool_cellStyle(xml2::xml_attr(node, "applyFill"))) {
    out$fill <- fills[[as.integer(xml2::xml_attr(node, "fillId")) + 1]]
  }
  if (bool_cellStyle(xml2::xml_attr(node, "applyBorder"))) {
    out$border <- borders[[as.integer(xml2::xml_attr(node, "borderId")) + 1]]
  }
  if (bool_cellStyle(xml2::xml_attr(node, "applyAlignment"))) {
    out$alignment = list(horizontal = xml2::xml_attr(alignment, "horizontal"),
                         vertical = xml2::xml_attr(alignment, "vertical"),
                         indent = xml2::xml_attr(alignment, "indent"),
                         wrap_text = xml2::xml_attr(alignment, "wrapText"),
                         justify_last_line = xml2::xml_attr(alignment, "justifyLastLine"),
                         shrink_to_fit = xml2::xml_attr(alignment, "shrinkToFit"),
                         reading_order = xml2::xml_attr(alignment, "readingOrder")
                         )
  }
  if (bool_cellStyle(xml2::xml_attr(node, "applyProtection"))) {
    out$protection = list(locked = xml2::xml_attr(protection, "locked"),
                          hidden = xml2::xml_attr(protection, "hidden"))
  }
  out
}

cellXf <- function(node, cellStyleXfs, numFmts, fonts, fills, borders) {
  alignment <- xml2::xml_find_first(node, "alignment")
  protection <- xml2::xml_find_first(node, "protection")
  out <- cellStyleXfs[[as.integer(xml2::xml_attr(node, "xfId")) + 1]] # Start fromo style, then modify
  bool_cell <- function(attribute) { # Whether to modify something
    bool <- as.integer(attribute)
    bool & (!is.na(bool))
  }
  if (bool_cell(xml2::xml_attr(node, "applyNumberFormat"))) {
    out$number_format <- numFmts[as.integer(xml2::xml_attr(node, "numFmtId")) + 1]
  }
  if (bool_cell(xml2::xml_attr(node, "applyFont"))) {
    out$font <- fonts[[as.integer(xml2::xml_attr(node, "fontId")) + 1]]
  }
  if (bool_cell(xml2::xml_attr(node, "applyFill"))) {
    out$fill <- fills[[as.integer(xml2::xml_attr(node, "fillId")) + 1]]
  }
  if (bool_cell(xml2::xml_attr(node, "applyBorder"))) {
    out$border <- borders[[as.integer(xml2::xml_attr(node, "borderId")) + 1]]
  }
  if (bool_cell(xml2::xml_attr(node, "applyAlignment"))) {
    out$alignment = list(horizontal = xml2::xml_attr(alignment, "horizontal"),
                         vertical = xml2::xml_attr(alignment, "vertical"),
                         indent = xml2::xml_attr(alignment, "indent"),
                         wrap_text = xml2::xml_attr(alignment, "wrapText"),
                         justify_last_line = xml2::xml_attr(alignment, "justifyLastLine"),
                         shrink_to_fit = xml2::xml_attr(alignment, "shrinkToFit"),
                         reading_order = xml2::xml_attr(alignment, "readingOrder")
                         )
  }
  if (bool_cell(xml2::xml_attr(node, "applyProtection"))) {
    out$protection = list(locked = xml2::xml_attr(protection, "locked"),
                          hidden = xml2::xml_attr(protection, "hidden"))
  }
  out
}

transpose_formats <- function(formats, colour_scheme) {
  formats %>%
    # Look up colours in index/theme
    purrr::map(~ purrr::map_at(.x, "font", ~ purrr::map_at(.x, "color", decode_color,
                                               colour_scheme = colour_scheme))) %>%
    purrr::map(~ purrr::map_at(.x, "fill",
                               ~ purrr::map_at(.x, c("bgColor", "fgColor"), decode_color,
                                               colour_scheme = colour_scheme))) %>%
    purrr::map(~ purrr::map_at(.x, "border",
                               ~ purrr::map(.x,
                                            ~ purrr::map_at(.x, "color", decode_color,
                                                            colour_scheme = colour_scheme)))) %>%
    # Transpose
    purrr::transpose() %>%
    purrr::map(purrr::transpose) %>%
    purrr::map_at("number_format", purrr::flatten) %>%
    purrr::map_at("font", ~ purrr::map_at(.x, "color", purrr::transpose)) %>%
    purrr::map_at("fill", ~ purrr::map_at(.x, c("bgColor", "fgColor"), purrr::transpose)) %>%
    purrr::map_at("border", ~ purrr::map(.x, purrr::transpose)) %>%
    purrr::map_at("border", ~ purrr::map(.x, ~ purrr::map_at(.x, "color", purrr::transpose))) %>%
    # Convert leaf lists to vectors
    purrr::map_at("number_format", purrr::as_vector) %>%
    purrr::map_at("font", ~ purrr::map_at(.x,
                            c("bold", "italic", "underline", "strike",
                              "size", "name", "family", "scheme"),
                            purrr::as_vector)) %>%
    purrr::map_at("font", ~ purrr::map_at(.x, "color", ~ purrr::map(.x, purrr::as_vector))) %>%
    purrr::map_at("fill", ~ purrr::map_at(.x, "patternType", purrr::as_vector)) %>%
    purrr::map_at("fill", ~ purrr::map_at(.x, c("bgColor", "fgColor"), ~ purrr::map(.x, purrr::as_vector))) %>%
    purrr::map_at("border", ~ purrr::map(.x, ~ purrr::map_at(.x, "style", purrr::as_vector))) %>%
    purrr::map_at("border", ~ purrr::map(.x, ~ purrr::map_at(.x, "color", ~ purrr::map(.x, purrr::as_vector)))) %>%
    purrr::map_at("alignment", ~ purrr::map(.x, purrr::as_vector)) %>%
    purrr::map_at("protection", ~ purrr::map(.x, purrr::as_vector))
}

