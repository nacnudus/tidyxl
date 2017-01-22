#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxbook.h"
#include "styles.h"
#include "utils.h"

using namespace Rcpp;

xlsxbook::xlsxbook(const std::string& path): path_(path), styles_(path_) {
  std::string book = zip_buffer(path_, "xl/workbook.xml");

  rapidxml::xml_document<> xml;
  xml.parse<0>(&book[0]);

  rapidxml::xml_node<>* workbook = xml.first_node("workbook");
  rapidxml::xml_node<>* sheets = workbook->first_node("sheets");

  cacheDateOffset(workbook); // Must come before cacheSheets
  cacheSheets(sheets);
  cacheStrings();
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
}

void xlsxbook::cacheDateOffset(rapidxml::xml_node<>* workbook) {
  rapidxml::xml_node<>* workbookPr = workbook->first_node("workbookPr");
  if (workbookPr != NULL) {
    rapidxml::xml_attribute<>* date1904 = workbookPr->first_attribute("date1904");
    if (date1904 != NULL) {
      dateOffset_ = 25569;
      return;
    }
  }

  dateOffset_ = 24107;
}
