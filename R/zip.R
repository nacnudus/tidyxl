# This code was copied from the R package 'readxl' at 
# https://cran.r-project.org/web/packages/readxl/index.html
# on 4 July 2016.
# It was written by Hadley Wickham hadley@rstudio.com.
# The copyright holder is the RStudio company.
# It is licensed under GPL-3.

# Called only from C++ code, but currently needs to be implemented in R.
zip_buffer <- function(zip_path, file_path) {
  files <- utils::unzip(zip_path, list = TRUE)

  indx <- match(file_path, files$Name)
  if (is.na(indx)) {
    stop("Couldn't find '", file_path, "' in '", zip_path, "'", call. = FALSE)
  }

  size <- files$Length[indx]

  con <- unz(zip_path, file_path, open = "rb")
  on.exit(close(con), add = TRUE)
  readBin(con, raw(), n = size)
}

zip_has_file <- function(zip_path, file_path) {
  file_path %in% utils::unzip(zip_path, list = TRUE)$Name
}

