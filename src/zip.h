// This code was adapted from the R package 'readxl' at
// https://cran.r-project.org/web/packages/readxl/index.html
// on 4 July 2016.
// It was written by Hadley Wickham hadley@rstudio.com.
// The copyright holder is the RStudio company.
// It is licensed under GPL-3.

#ifndef TIDYXL_ZIP_
#define TIDYXL_ZIP_

#include "rapidxml.h"

std::string zip_buffer(const std::string& zip_path, const std::string& file_path);
bool zip_has_file(const std::string& zip_path, const std::string& file_path);
std::string extdata();

#endif

