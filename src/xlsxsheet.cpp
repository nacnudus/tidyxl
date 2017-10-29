#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxcell.h"
#include "xlsxsheet.h"
#include "xlsxbook.h"
#include "string.h"
#include "xml_value.h"

using namespace Rcpp;

xlsxsheet::xlsxsheet(
    const std::string& name,
    std::string& sheet_xml,
    xlsxbook& book,
    String comments_path):
  name_(name),
  book_(book) {
  rapidxml::xml_document<> xml;
  xml.parse<rapidxml::parse_non_destructive>(&sheet_xml[0]);

  rapidxml::xml_node<>* worksheet = xml.first_node("worksheet");
  rapidxml::xml_node<>* sheetData = worksheet->first_node("sheetData");

  cacheDefaultRowColDims(worksheet);
  cacheColWidths(worksheet);
  cacheComments(comments_path);
  cacheCellcount(sheetData);
}

void xlsxsheet::cacheDefaultRowColDims(rapidxml::xml_node<>* worksheet) {
  rapidxml::xml_node<>* sheetFormatPr_ = worksheet->first_node("sheetFormatPr");

  if (sheetFormatPr_ != NULL) {
    double_attr(defaultRowHeight_, sheetFormatPr_, "defaultRowHeight", 15);

    // If defaultColWidth not given, ECMA says you can work it out based on
    // baseColWidth, but that isn't necessarily given either, and the formula
    // is wrong because the reality is so complicated, see
    // https://support.microsoft.com/en-gb/kb/214123.
    double_attr(defaultColWidth_, sheetFormatPr_, "defaultColWidth", 8.38);
  }
}

void xlsxsheet::cacheColWidths(rapidxml::xml_node<>* worksheet) {
  // Having done cacheDefaultRowColDims(), initilize vector to default width,
  // then update with custom widths.  The number of columns might be available
  // by parsing <dimension><ref>, but if not then only by parsing the address of
  // all the cells.  I think it's better just to use the maximum possible number
  // of columns, 16834.

  colWidths_.assign(16384, defaultColWidth_);

  rapidxml::xml_node<>* cols = worksheet->first_node("cols");
  if (cols == NULL)
    return; // No custom widths

  for (rapidxml::xml_node<>* col = cols->first_node("col");
      col; col = col->next_sibling("col")) {

    // <col> applies to columns from a min to a max, which must be iterated over
    unsigned int min  = std::stol(col->first_attribute("min")->value());
    unsigned int max  = std::stol(col->first_attribute("max")->value());
    double width = std::stod(col->first_attribute("width")->value());

    for (unsigned int column = min; column <= max; ++column)
      colWidths_[column - 1] = width;
  }
}

unsigned long long int xlsxsheet::cacheCellcount(rapidxml::xml_node<>* sheetData) {
  // Iterate over all rows and cells to count.  The 'dimension' tag is no use
  // here because it describes a rectangle of cells, many of which may be blank.
  unsigned long long int cellcount = 0;
  unsigned long long int commentcount = 0; // no. of matching comments
  rapidxml::xml_attribute<>* r;
  std::map<std::string, std::string>::iterator comment;
  for (rapidxml::xml_node<>* row = sheetData->first_node("row");
      row; row = row->next_sibling("row")) {
    for (rapidxml::xml_node<>* c = row->first_node("c");
      c; c = c->next_sibling("c")) {
      // Check for a matching comment.
      // Comments on blank cells don't have a matching cell, so a blank cell
      // must be created for them.  Such additional cells must be added to the
      // actual cellcount, to initialize the returned vectors to the correct
      // length.
      r = c->first_attribute("r");
      if (r == NULL) // check once in whole program
        stop("Invalid row or cell: lacks 'r' attribute");
      comment = comments_.find(std::string(r->value(), r->value_size()));
      if(comment != comments_.end()) {
        ++commentcount;
      }
      ++cellcount;
      if ((cellcount + 1) % 1000 == 0) {
        checkUserInterrupt();
      }
    }
  }
  cellcount_ = cellcount + (comments_.size() - commentcount);
  return(cellcount_);
}

void xlsxsheet::cacheComments(String comments_path) {
  // Having constructed the map, they will each be deleted when they are matched
  // to a cell.  That will leave only those comments that are on empty cells.
  // Those are then appended as empty cells with comments.
  if (comments_path != NA_STRING) {
    std::string comments_file = zip_buffer(book_.path_, comments_path);
    rapidxml::xml_document<> xml;
    xml.parse<rapidxml::parse_strip_xml_namespaces>(&comments_file[0]);

    // Iterate over the comments to store the ref and text
    rapidxml::xml_node<>* comments = xml.first_node("comments");
    rapidxml::xml_node<>* commentList = comments->first_node("commentList");
    for (rapidxml::xml_node<>* comment = commentList->first_node();
        comment; comment = comment->next_sibling()) {
      rapidxml::xml_attribute<>* ref = comment->first_attribute("ref");
      std::string ref_string(ref->value(), ref->value_size());
      rapidxml::xml_node<>* r = comment->first_node();
      // Get the inline string
      std::string inlineString;
      parseString(r, inlineString); // value is modified in place
      comments_[ref_string] = inlineString;
    }
  }
}

void xlsxsheet::parseSheetData(
    rapidxml::xml_node<>* sheetData,
    unsigned long long int& i) {
  // Iterate through rows and cells in sheetData.  Cell elements are children
  // of row elements.  Columns are described elswhere in cols->col.
  rowHeights_.assign(1048576, defaultRowHeight_); // cache rowHeight while here
  unsigned long int rowNumber;
  for (rapidxml::xml_node<>* row = sheetData->first_node();
      row; row = row->next_sibling()) {
    rapidxml::xml_attribute<>* r = row->first_attribute("r");
    if (r == NULL)
      stop("Invalid row or cell: lacks 'r' attribute");
    rowNumber = std::stod(r->value());
    // Check for custom row height
    double rowHeight = defaultRowHeight_;
    rapidxml::xml_attribute<>* ht = row->first_attribute("ht");
    if (ht != NULL) {
      rowHeight = std::stod(ht->value());
      rowHeights_[rowNumber - 1] = rowHeight;
    }

    for (rapidxml::xml_node<>* c = row->first_node();
        c; c = c->next_sibling()) {
      xlsxcell cell(c, this, book_, i);

      // Sheet name, row height and col width aren't really determined by the
      // cell, so they're done in this sheet instance
      book_.sheet_[i] = name_;
      book_.height_[i] = rowHeight;
      book_.width_[i] = colWidths_[book_.col_[i] - 1];

      ++i;
      if ((i + 1) % 1000 == 0)
        checkUserInterrupt();
    }
  }
}

void xlsxsheet::appendComments(unsigned long long int& i) {
  // Having constructed the comments_ map, they are each be deleted when they
  // are matched to a cell.  That leaves only those comments that are on empty
  // cells.  This code appends those remaining comments as empty cells.
  int col;
  int row;
  for(std::map<std::string, std::string>::iterator it = comments_.begin();
      it != comments_.end(); ++it) {
    // TODO: move address parsing to utils
    std::string address = it->first.c_str(); // we need this std::string in a moment
    // Iterate though the A1-style address string character by character
    col = 0;
    row = 0;
    for(std::string::const_iterator iter = address.begin();
        iter != address.end(); ++iter) {
      if (*iter >= '0' && *iter <= '9') { // If it's a number
        row = row * 10 + (*iter - '0'); // Then multiply existing row by 10 and add new number
      } else if (*iter >= 'A' && *iter <= 'Z') { // If it's a character
        col = 26 * col + (*iter - 'A' + 1); // Then do similarly with columns
      }
    }
    book_.sheet_[i] = name_;
    book_.address_[i] = address;
    book_.row_[i] = row;
    book_.col_[i] = col;
    book_.is_blank_[i] = true;
    book_.data_type_[i] = "blank";
    book_.error_[i] = NA_STRING;
    book_.logical_[i] = NA_LOGICAL;
    book_.numeric_[i] = NA_REAL;
    book_.date_[i] = NA_REAL;
    book_.character_[i] = NA_STRING;
    book_.formula_[i] = NA_STRING;
    book_.is_array_[i] = false;
    book_.formula_ref_[i] = NA_STRING;
    book_.formula_group_[i] = NA_INTEGER;
    book_.comment_[i] = it->second.c_str();
    book_.height_[i] = rowHeights_[row - 1];
    book_.width_[i] = colWidths_[col - 1];
    book_.style_format_[i] = "Normal";
    book_.local_format_id_[i] = 1;
    ++i;
  }
  // Iterate though the A1-style address string character by character
}

