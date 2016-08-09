#' Join a bag of data cells to some header, by proximity in a given direction, e.g. NNW is up and up-left.
#'
#' @description A bag of data cells is a data frame with at least the
#' columns 'row' and 'col', as well as any others that carry information about
#' the cells, e.g. their values.  Cells in a table are associated with header
#' cells by proximity.  Having collected header cells and data cells into
#' separate data frames, 'join_header' and the related functions 'NNW', 'ABOVE',
#' etc., join the values in the header cells to the data cells, choose the
#' nearest header to each cell, in a given direction.
#' @param bag Data frame. A bag of data cells including at least the columns
#' 'row' and 'column', which are numeric/integer vectors.
#' @param header Data frame. A bag of header cells, which currently must have
#' the columns 'row', 'col' and 'header', where 'header' is a column whose value
#' describes the data.  For example, for data cells recording a number of
#' animals, and a header describing the type of animal, then the 'header' column
#' might contain the values `c("dog", "cat", "mouse")`.
#' @param direction Character vector length 1. A compass direction to search for the nearest header.  See
#' 'details'.
#' @param colname Character vector length one. Column name to give the header
#' values once they are joined to the 'bag'.  Continuing the example, this might
#' be "animal_type".
#' @details Headers are associated with data by proximity in a given direction.
#' The directions are mapped to the points of the compass, where 'N' is north
#' (up), 'E' is east (right), and so on.  `join_header` finds the nearest header
#' to a given data cell in a given direction, and joins its value to the data
#' cell.  The most common directions to search are 'NNW' (for left-aligned
#' headers at the top of the table) and 'WNW' for top-aligned headers at the
#' side of the table.  There can be a tie in the directions 'ABOVE', 'BELOW',
#' 'LEFT' and 'RIGHT' (for headers that are not aligned to the edge of the data
#' cells that they refer to), and currently ties will cause the affected cells
#' not to be returned at all.  The full list of available directions is 'N',
#' 'E', 'S', 'W', 'NE', 'NE', 'SE', 'SW', 'NNW', 'NNE', 'ENE', 'ESE', 'SSE',
#' 'SSW', 'WSW', 'WNW'.  For convenience, these directions are provided as their
#' own functions, wrapping the concept of 'join_header'.
#' @name join_header
#' @export
join_header <- function(bag, header, direction, colname = "newheader") {
  # Extends an existing bag of cells along an axis up to a boundary, by row or
  # by column depending on the axis.
  # Bag may be ragged rows or ragged cols, but gaps will not be filled in.
  # Prototype extending right
  if (direction == "N") {
    N(bag, header, colname)
  } else if (direction == "W") {
    W(bag, header, colname)
  } else if (direction == "NNW") {
    NNW(bag, header, colname)
  } else if (direction == "WNW") {
    WNW(bag, header, colname)
  } else if (direction == "ABOVE") {
    ABOVE(bag, header, colname)
  } else if (direction == "LEFT") {
    LEFT(bag, header, colname)
  } else {stop("The direction ", direction, 
               ", is either not recognised or not yet supported.")}
}

#' @describeIn join_header Join nearest header in the 'N' direction.
#' @export
N <- function(bag, header, colname = "newheader") {
  # Join header to cells by proximity
  # First, the domain of each header
  domains <- 
    header %>%
    arrange(row) %>%
    # x1 and y1 are the cell itself, x2 is the same column as the cell itself
    mutate(x1 = as.numeric(col),
           y1 = as.numeric(row),
           x2 = as.numeric(col)) %>%
    # y2 goes up to just before the next header in any column
    group_by(row) %>%
    nest %>%
    mutate(y2 = lead(as.numeric(row) - 1, default = Inf)) %>%
    unnest %>%
    ungroup
  get_header(bag, domains, colname)
}

#' @describeIn join_header Join nearest header in the 'W' direction.
#' @export
W <- function(bag, header, colname = "newheader") {
  # Join header to cells by proximity
  # First, the domain of each header
  domains <- 
    header %>%
    arrange(col) %>%
    # x1 and y1 are the cell itself, y2 is the same row as the cell itself
    mutate(x1 = as.numeric(col),
           y1 = as.numeric(row),
           y2 = as.numeric(row)) %>%
    # x2 goes up to just before the next header in any row
    group_by(col) %>%
    nest %>%
    mutate(x2 = lead(as.numeric(col) - 1, default = Inf)) %>%
    unnest %>%
    ungroup
  get_header(bag, domains, colname)
}

#' @describeIn join_header Join nearest header in the 'NNW' direction.
#' @export
NNW <- function(bag, header, colname = "newheader") {
  # Join header to cells by proximity
  # First, the domain of each header
  domains <- 
    header %>%
    arrange(row, col) %>%
    # x1 and y1 are the cell itself
    mutate(x1 = as.numeric(col),
           y1 = as.numeric(row)) %>%
    # x2 goes up to just before the next header in the row
    group_by(row) %>%
    mutate(x2 = lead(as.numeric(col) - 1, default = Inf)) %>%
    # y2 goes up to just before the next header in any column
    nest %>%
    mutate(y2 = lead(as.numeric(row) - 1, default = Inf)) %>%
    unnest %>%
    ungroup
  get_header(bag, domains, colname)
}

#' @describeIn join_header Join nearest header in the 'WNW' direction.
#' @export
WNW <- function(bag, header, colname = "newheader") {
  # Join header to cells by proximity
  # First, the domain of each header
  domains <- 
    header %>%
    arrange(col, row) %>%
    # x1 and y1 are the cell itself
    mutate(x1 = as.numeric(col),
           y1 = as.numeric(row)) %>%
    # y2 goes up to just before the next header in the column
    group_by(col) %>%
    mutate(y2 = lead(as.numeric(row) - 1, default = Inf)) %>%
    # x2 goes up to just before the next header in any row
    nest %>%
    mutate(x2 = lead(as.numeric(col) - 1, default = Inf)) %>%
    unnest %>%
    ungroup
  get_header(bag, domains, colname)
}

#' @describeIn join_header Join nearest header in the 'ABOVE' direction.
#' @export
ABOVE <- function(bag, header, colname = "newheader") {
  # Join header to cells by proximity
  # First, the domain of each header
  domains <- 
    header %>%
    arrange(row, col) %>%
    # y1 is the cell itself
    mutate(y1 = as.numeric(row)) %>%
    # x1 and x2 are half-way (rounded down) from the cell to headers either
    # side in the same row
    group_by(row) %>%
    mutate(x1 = floor((col + lag(as.numeric(col), default = -Inf) + 2)/2),
           x2 = ceiling((col + lead(as.numeric(col), default = Inf) - 2)/2)) %>%
    # y2 goes up to just before the next header in any column
    nest %>%
    mutate(y2 = lead(as.numeric(row) - 1, default = Inf)) %>%
    unnest %>%
    ungroup
  get_header(bag, domains, colname)
}

#' @describeIn join_header Join nearest header in the 'LEFT' direction.
#' @export
LEFT <- function(bag, header, colname = "newheader") {
  # Join header to cells by proximity
  # First, the domain of each header
  domains <- 
    header %>%
    arrange(col, row) %>%
    # x1 is the cell itself
    mutate(x1 = as.numeric(col)) %>%
    # r1 and y2 are half-way (rounded down) from the cell to headers either
    # side in the same column
    group_by(col) %>%
    mutate(y1 = floor((row + lag(as.numeric(row), default = -Inf) + 2)/2),
           y2 = ceiling((row + lead(as.numeric(row), default = Inf) - 2)/2)) %>%
    # x2 goes up to just before the next header in any row
    nest %>%
    mutate(x2 = lead(as.numeric(col) - 1, default = Inf)) %>%
    unnest %>%
    ungroup
  get_header(bag, domains, colname)
}

get_header <- function(bag, domains, colname) {
  # Use data.table non-equi join to join header with cells.
  bag <-
    bag %>%
    mutate(row = as.numeric(row), col = as.numeric(col)) # joins with columns that are numeric to allow NA
  bag <- data.table(bag)         # Must be done without %>%
  domains <- data.table(domains) # Must be done without %>%
  joined <- 
    bag[domains,
          .(row = x.row, col = x.col, header = i.header),
          on = .(row >= y1, row <= y2, col >= x1, col <= x2)] %>%
    rename_(.dots = setNames(list(~header), colname))
  # Finally, join back on the cells (this step is necessary because data.table
  # returns weird columns from a non-equi join)
  bag %>%
    inner_join(joined, by = c("row", "col")) %>%
    tbl_df
}
