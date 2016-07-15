#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "rapidxml_print.h"
#include "xlsxbook.h"

using namespace Rcpp;

xlsxbook::xlsxbook(const std::string& path) {
  std::string book_ = zip_buffer(path, "xl/workbook.xml");
  bookXml_.parse<0>(&book_[0]);

  rootNode_ = bookXml_.first_node("workbook");
  if (rootNode_ == NULL)
    stop("Invalid workbook xml (no <workbook>)");

  rapidxml::xml_node<>* sheets_;
  sheets_ = rootNode_->first_node("sheets");
  if (sheets_ == NULL)
    stop("Invalid workbook xml (no <sheets>)");

  cacheSheets(sheets_);
}

List xlsxbook::information() {
  /* // Print a node to a string -- could be good for inline-formatted cell text */
  /* std::string s; */
  /* print(std::back_inserter(s), *sheets_, 0); */

  NumericVector x = NumericVector::create(1);
  return List::create(Named("x") = sheets_);
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

const std::vector<std::string>& xlsxbook::sheets() {
  return sheets_;
}
