#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "rapidxml_print.h"
#include "xlsxsheet.h"
#include "xlsxbook.h"

using namespace Rcpp;

xlsxsheet::xlsxsheet(
    const std::string& bookPath,
    const int& i, xlsxbook& book) {
  std::string sheetPath = tfm::format("xl/worksheets/sheet%i.xml", i + 1);
  std::string sheet_ = zip_buffer(bookPath, sheetPath);
  xml_.parse<0>(&sheet_[0]);

  worksheet_ = xml_.first_node("worksheet");
  if (worksheet_ == NULL)
    stop("Invalid sheet xml (no <worksheet>)");

  sheetData_ = worksheet_->first_node("sheetData");
  if (sheetData_ == NULL)
    stop("Invalid sheet xml (no <sheetData>)");

  // Look up name among worksheets in book
  name_ = book.sheets()[i];
}

List xlsxsheet::information() {
  // Print a node to a string -- could be good for inline-formatted cell text
  std::string s;
  print(std::back_inserter(s), *sheetData_, 0);

  NumericVector x = NumericVector::create(1);
  return List::create(Named(name_) = s);
}
