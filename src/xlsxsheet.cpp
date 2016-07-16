#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxcell.h"
#include "xlsxsheet.h"
#include "xlsxbook.h"
#include "utils.h"

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

  cacheDefaultRowColDims();
  cacheCellcount();
  initializeColumns();
  parseSheetData();
}

List xlsxsheet::information() {
  // Returns a nested data frame of everything, the data frame itself wrapped in
  // a list.
  List data = List::create(
      address_,
      row_,
      col_,
      content_,
      type_,
      height_,
      width_);
  CharacterVector colnames = CharacterVector::create(
      "address",
      "row",
      "col",
      "content",
      "type",
      "height",
      "width");
  makeDataFrame(data, colnames);
  return data;
}

void xlsxsheet::cacheDefaultRowColDims() {
  rapidxml::xml_node<>* sheetFormatPr_ = worksheet_->first_node("sheetFormatPr");

  if (sheetFormatPr_ != NULL) {
    // Don't use utils::getAttributeValueDouble because it might overwrite the 
    // default value with NA_REAL.
    rapidxml::xml_attribute<>* defaultRowHeight = 
      sheetFormatPr_->first_attribute("defaultRowHeight");
    if (defaultRowHeight != NULL)
      defaultRowHeight_ = atof(defaultRowHeight->value());

    rapidxml::xml_attribute<>*defaultColWidth = 
      sheetFormatPr_->first_attribute("defaultColWidth");
    if (defaultColWidth != NULL)
      defaultColWidth_ = atof(defaultColWidth->value());
      // If defaultColWidth not given, ECMA says you can work it out based on
      // baseColWidth, but that isn't necessarily given either, and the formula
      // is wrong because the reality is so complicated, see 
      // https://support.microsoft.com/en-gb/kb/214123.
  }
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

void xlsxsheet::initializeColumns() {
  // Having counted the cells, make columns of that length
  address_   = CharacterVector(cellcount_, NA_STRING);
  row_       = IntegerVector(cellcount_,   NA_INTEGER);
  col_       = IntegerVector(cellcount_,   NA_INTEGER);
  content_   = CharacterVector(cellcount_, NA_STRING);
  value_     = List(cellcount_);
  type_      = CharacterVector(cellcount_, NA_STRING);
  logical_   = LogicalVector(cellcount_,   NA_LOGICAL);
  numeric_   = NumericVector(cellcount_,   NA_REAL);
  date_      = NumericVector(cellcount_,   NA_REAL);
    date_.attr("class") = Rcpp::CharacterVector::create("POSIXct", "POSIXt");
    date_.attr("tzone") = "UTC";
  character_ = CharacterVector(cellcount_, NA_STRING);
  error_     = CharacterVector(cellcount_, NA_STRING);
  height_    = NumericVector(cellcount_,   NA_REAL);
  width_     = NumericVector(cellcount_,   NA_REAL);
}

void xlsxsheet::parseSheetData() {
  // Iterate through rows and cells in sheetData_.  Cell elements are children
  // of row elements.  Columns are described elswhere in cols->col.

  unsigned long long int i = 0; // counter for checkUserInterrupt
  for (rapidxml::xml_node<>* row = sheetData_->first_node("row");
      row; row = row->next_sibling("row")) {
    if ((i + 1) % 1000 == 0)
      Rcpp::checkUserInterrupt();

    // Check for custom row height
    double rowheight = defaultRowHeight_;
    rapidxml::xml_attribute<>* ht = row->first_attribute("ht");
    if (ht != NULL) 
      rowheight = atof(ht->value());

    // Col widths must be looked up among cols->col, which combine multiple cols
    // of the same properties.  Not yet implemented.

    for (rapidxml::xml_node<>* c = row->first_node("c");
        c; c = c->next_sibling("c")) {
      xlsxcell cell(c, rowheight);

      address_[i] = cell.address();
      row_[i] = cell.row();
      col_[i] = cell.col();
      content_[i] = cell.content();
      type_[i] = cell.type();
      height_[i] = cell.height();
      width_[i] = cell.width();

      ++i;
    }
  }
}
