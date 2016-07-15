#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxcell.h"
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

  cacheCellcount();
  parseSheetData();
}

List xlsxsheet::information() {
  NumericVector x = NumericVector::create(1);
  return List::create(Named(name_) = std::to_string(cellcount_));
}

void xlsxsheet::cacheCellcount() {
  // Iterate over all rows and cells to count
  cellcount_ = 0;
  for (rapidxml::xml_node<>* row = sheetData_->first_node("row");
      row; row = row->next_sibling("row")) {
    for (rapidxml::xml_node<>* c = row->first_node("c");
        c; c = c->next_sibling("c")) {
      ++cellcount_;
    }
  }
}

void xlsxsheet::parseSheetData() {
  // Iterate through rows and cells in sheetData_
  unsigned long long int i = 0; // counter for checkUserInterrupt
  for (rapidxml::xml_node<>* row = sheetData_->first_node("row");
      row; row = row->next_sibling("row")) {
    if ((i + 1) % 1000 == 0)
      Rcpp::checkUserInterrupt();
    for (rapidxml::xml_node<>* c = row->first_node("c");
        c; c = c->next_sibling("c")) {
      xlsxcell cell(c);
    }
    ++i;
  }
}
