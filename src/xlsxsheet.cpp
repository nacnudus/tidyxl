#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "rapidxml_print.h"
#include "xlsxsheet.h"

std::string sheet_;
rapidxml::xml_document<> sheetXml_;
rapidxml::xml_node<>* rootNode_;
rapidxml::xml_node<>* sheetData_;

xlsxsheet::xlsxsheet(const std::string& bookPath, const int& i) {
  std::string sheetPath = tfm::format("xl/worksheets/sheet%i.xml", i + 1);
  sheet_ = zip_buffer(bookPath, sheetPath);
  sheetXml_.parse<0>(&sheet_[0]);

  rootNode_ = sheetXml_.first_node("worksheet");
  if (rootNode_ == NULL)
    Rcpp::stop("Invalid sheet xml (no <worksheet>)");

  sheetData_ = rootNode_->first_node("sheetData");
  if (sheetData_ == NULL)
    Rcpp::stop("Invalid sheet xml (no <sheetData>)");
}

Rcpp::List xlsxsheet::information(const int& i) {
  // Print a node to a string -- could be good for inline-formatted cell text
  std::string s;
  print(std::back_inserter(s), *sheetData_, 0);

  Rcpp::NumericVector x = Rcpp::NumericVector::create(i);
  return Rcpp::List::create(Rcpp::Named("x") = s);
}
