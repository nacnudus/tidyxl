#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxbook.h"
#include "utils.h"

using namespace Rcpp;

xlsxbook::xlsxbook(const std::string& path): path_(path) {
  std::string book = zip_buffer(path_, "xl/workbook.xml");

  rapidxml::xml_document<> xml;
  xml.parse<0>(&book[0]);

  rapidxml::xml_node<>* workbook = xml.first_node("workbook");
  if (workbook == NULL)
    stop("Invalid workbook xml (no <workbook>)");

  rapidxml::xml_node<>* sheets = workbook->first_node("sheets");
  if (sheets == NULL)
    stop("Invalid workbook xml (no <sheets>)");

  cacheSheets(sheets);
  cacheStrings();
  cacheCellXfsXfId();
}

std::string xlsxbook::path() {
  return path_;
}

std::vector<std::string>& xlsxbook::sheets() {
  return sheets_;
}

std::vector<std::string>& xlsxbook::strings() {
  return strings_;
}

unsigned long int xlsxbook::strings_size() {
  return strings_size_;
}

std::vector<int>& xlsxbook::cellXfs_xfId() {
  return cellXfs_xfId_;
}

void xlsxbook::cacheSheets(rapidxml::xml_node<>* sheets) {
  // ECMA specifies no maximum number of sheets
  // Most often it will be 3, but two resizes won't matter much, so I don't
  // bother reserving.
  // http://stackoverflow.com/a/7397862/937932 recommends not reserving anyway.

  for (rapidxml::xml_node<>* sheet = sheets->first_node();
      sheet; sheet = sheet->next_sibling()) {
    std::string name = sheet->first_attribute("name")->value();
    sheets_.push_back(name);
  }
}

// Based on hadley/readxl
void xlsxbook::cacheStrings() {
  if (!zip_has_file(path_, "xl/sharedStrings.xml"))
    return;

  std::string xml = zip_buffer(path_, "xl/sharedStrings.xml");
  rapidxml::xml_document<> sharedStrings;
  sharedStrings.parse<0>(&xml[0]);

  rapidxml::xml_node<>* sst = sharedStrings.first_node("sst");
  if (sst == NULL)
    return;

  rapidxml::xml_attribute<>* uniqueCount = sst->first_attribute("uniqueCount");
  if (uniqueCount != NULL) {
    unsigned long int n = strtol(uniqueCount->value(), NULL, 10);
    strings_.reserve(n);
  }

  // 18.4.8 si (String Item) [p1725]
  for (rapidxml::xml_node<>* string = sst->first_node();
       string; string = string->next_sibling()) {
    std::string out;
    parseString(string, out);    // missing strings are treated as empty ""
    strings_.push_back(out);
  }

  strings_size_ = strings_.size();
}

// Create a vector of the theme id of each cell style.
// Cells have cell styles, which link to theme styles.  The cell style id is
// given in the cell itself (in the worksheet xml), but the linked theme id is
// only given in the cell's style record, so go via that.
void xlsxbook::cacheCellXfsXfId() {
  if (!zip_has_file(path_, "xl/styles.xml"))
    return;

  std::string xml = zip_buffer(path_, "xl/styles.xml");
  rapidxml::xml_document<> styles;
  styles.parse<0>(&xml[0]);

  rapidxml::xml_node<>* styleSheet = styles.first_node("styleSheet");
  if (styleSheet == NULL)
    return;

  rapidxml::xml_node<>* cellXfs = styleSheet->first_node("cellXfs");
  if (cellXfs == NULL)
    return;

  rapidxml::xml_attribute<>* count = cellXfs->first_attribute("count");
  if (count != NULL) {
    int n = atoi(count->value());
    int i = 0;
    cellXfs_xfId_.reserve(n);
    for (rapidxml::xml_node<>* xf = cellXfs->first_node();
         xf; xf = xf->next_sibling()) {
      rapidxml::xml_attribute<>* xfId = xf->first_attribute("xfId");
      if (xfId == NULL) {
        // This happened in ./tests/testthat/iris-google-doc.xlsx a
        // google-sheets-exported file, where there's no theme1.xml to index
        // into anyway.
        cellXfs_xfId_[i] = NA_INTEGER;
      } else {
        cellXfs_xfId_[i] = atoi(xfId->value()) + 1;
      }
      i++;
    }
  }
}
