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

  cacheDateOffset(workbook); // Must come before cacheSheets
  cacheStrings();
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
      std::string is1904 = date1904->value();
      if ((is1904 == "1") || (is1904 == "true")) {
        dateSystem_ = 1904;
        dateOffset_ = 24107;
        return;
      }
    }
  }

  dateSystem_ = 1900;
  dateOffset_ = 25569;
}
