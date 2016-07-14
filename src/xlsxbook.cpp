#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "rapidxml_print.h"
#include "xlsxbook.h"

xlsxbook::xlsxbook(const std::string& bookPath) {
  std::string book_ = zip_buffer(bookPath, "xl/workbook.xml");
  bookXml_.parse<0>(&book_[0]);

  rootNode_ = bookXml_.first_node("workbook");
  if (rootNode_ == NULL)
    Rcpp::stop("Invalid workbook xml (no <workbook>)");

  sheets_ = rootNode_->first_node("sheets");
  if (sheets_ == NULL)
    Rcpp::stop("Invalid workbook xml (no <sheets>)");
}

Rcpp::List xlsxbook::information() {
  // Print a node to a string -- could be good for inline-formatted cell text
  std::string s;
  print(std::back_inserter(s), *sheets_, 0);

  Rcpp::NumericVector x = Rcpp::NumericVector::create(1);
  return Rcpp::List::create(Rcpp::Named("x") = s);
}
