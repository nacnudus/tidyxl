#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxcell.h"
#include "xlsxsheet.h"
#include "xlsxbook.h"
#include "utils.h"

using namespace Rcpp;

xlsxsheet::xlsxsheet(const int& sheetindex, xlsxbook& book): book_(book) {
  std::string sheetpath = tfm::format("xl/worksheets/sheet%i.xml", sheetindex);
  std::string sheet_ = zip_buffer(book_.path(), sheetpath);
  xml_.parse<0>(&sheet_[0]);

  worksheet_ = xml_.first_node("worksheet");
  if (worksheet_ == NULL)
    stop("Invalid sheet xml (no <worksheet>)");

  sheetData_ = worksheet_->first_node("sheetData");
  if (sheetData_ == NULL)
    stop("Invalid sheet xml (no <sheetData>)");

  // Look up name among worksheets in book
  name_ = book_.sheets()[sheetindex - 1];

  defaultRowHeight_ = 15;
  double defaultColWidth_ = 8.47;

  cacheDefaultRowColDims();
  cacheColWidths();
  cacheCellcount();
  initializeColumns();
  parseSheetData();
}

List& xlsxsheet::information() {
  // Returns a nested data frame of everything, the data frame itself wrapped in
  // a list.
  information_ = List::create(
      address_,
      row_,
      col_,
      content_,
      formula_,
      formula_type_,
      formula_ref_,
      formula_group_,
      type_,
      character_,
      height_,
      width_,
      style_format_id_,
      local_format_id_);
  CharacterVector colnames = CharacterVector::create(
      "address",
      "row",
      "col",
      "content",
      "formula",
      "formula_type",
      "formula_ref",
      "formula_group",
      "type",
      "character",
      "height",
      "width",
      "style_format_id",
      "local_format_id");
  makeDataFrame(information_, colnames);
  return information_;
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

void xlsxsheet::cacheColWidths() {
  // Having done cacheDefaultRowColDims(), initilize vector to default width,
  // then update with custom widths.  The number of columns might be available
  // by parsing <dimension><ref>, but if not then only by parsing the address of
  // all the cells.  I think it's better just to use the maximum possible number
  // of columns, 16834.

  colWidths_.assign(16384, defaultColWidth_);

  rapidxml::xml_node<>* cols = worksheet_->first_node("cols");
  if (cols == NULL)
    return; // No custom widths

  for (rapidxml::xml_node<>* col = cols->first_node("col");
      col; col = col->next_sibling("col")) {

    // <col> applies to columns from a min to a max, which must be iterated over
    int min  = atoi(col->first_attribute("min")->value());
    int max  = atoi(col->first_attribute("max")->value());
    double width = atof(col->first_attribute("width")->value());

    for (int column = min; column <= max; ++column)
      colWidths_[column - 1] = width;
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
  // Having done cacheCellcount(), make columns of that length
  address_         = CharacterVector(cellcount_, NA_STRING);
  row_             = IntegerVector(cellcount_,   NA_INTEGER);
  col_             = IntegerVector(cellcount_,   NA_INTEGER);
  content_         = CharacterVector(cellcount_, NA_STRING);
  formula_         = CharacterVector(cellcount_, NA_STRING);
  formula_type_    = CharacterVector(cellcount_, NA_STRING);
  formula_ref_     = CharacterVector(cellcount_, NA_STRING);
  formula_group_   = IntegerVector(cellcount_,   NA_INTEGER);
  value_           = List(cellcount_);
  type_            = CharacterVector(cellcount_, NA_STRING);
  logical_         = LogicalVector(cellcount_,   NA_LOGICAL);
  numeric_         = NumericVector(cellcount_,   NA_REAL);
  date_            = NumericVector(cellcount_,   NA_REAL);
    date_.attr("class") = Rcpp::CharacterVector::create("POSIXct", "POSIXt");
    date_.attr("tzone") = "UTC";
  character_       = CharacterVector(cellcount_, NA_STRING);
  error_           = CharacterVector(cellcount_, NA_STRING);
  height_          = NumericVector(cellcount_,   NA_REAL);
  width_           = NumericVector(cellcount_,   NA_REAL);
  style_format_id_ = IntegerVector(cellcount_,   NA_INTEGER);
  local_format_id_ = IntegerVector(cellcount_,   NA_INTEGER);
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
    double rowHeight = defaultRowHeight_;
    rapidxml::xml_attribute<>* ht = row->first_attribute("ht");
    if (ht != NULL) 
      rowHeight = atof(ht->value());

    // Col widths are looked up among <cols><col>, the passed by reference to
    // the cell, which looks up its own column

    for (rapidxml::xml_node<>* c = row->first_node("c");
        c; c = c->next_sibling("c")) {
      xlsxcell cell(c, rowHeight, colWidths_, book_);

      address_[i] = cell.address();
      row_[i] = cell.row();
      col_[i] = cell.col();
      content_[i] = cell.content();
      formula_[i] = cell.formula();
      formula_type_[i] = cell.formula_type();
      formula_ref_[i] = cell.formula_ref();
      formula_group_[i] = cell.formula_group();
      type_[i] = cell.type();
      character_[i] = cell.character();
      height_[i] = cell.height();
      width_[i] = cell.width();
      style_format_id_[i] = cell.style_format_id();
      local_format_id_[i] = cell.local_format_id();

      ++i;
    }
  }
}
