#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxbook.h"
#include "utils.h"

using namespace Rcpp;

xlsxbook::xlsxbook(const std::string& path): path_(path) {
  std::string book_ = zip_buffer(path_, "xl/workbook.xml");
  xml_.parse<0>(&book_[0]);

  workbook_ = xml_.first_node("workbook");
  if (workbook_ == NULL)
    stop("Invalid workbook xml (no <workbook>)");

  rapidxml::xml_node<>* sheets_;
  sheets_ = workbook_->first_node("sheets");
  if (sheets_ == NULL)
    stop("Invalid workbook xml (no <sheets>)");

  cacheSheets(sheets_);
  cacheStrings();
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

  rapidxml::xml_attribute<>* count = sst->first_attribute("count");
  if (count != NULL) {
    int n = atoi(count->value());
    strings_.reserve(n);
  }

  // 18.4.8 si (String Item) [p1725]
  for (rapidxml::xml_node<>* string = sst->first_node();
       string; string = string->next_sibling()) {
    std::string out;
    parseString(string, out);    // missing strings are treated as empty ""
    strings_.push_back(out);
  }
}
