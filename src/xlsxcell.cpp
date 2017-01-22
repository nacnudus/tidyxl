#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"
#include "xlsxcell.h"
#include "xlsxsheet.h"
#include "utils.h"

using namespace Rcpp;

xlsxcell::xlsxcell(
    rapidxml::xml_node<>* cell,
    xlsxsheet* sheet,
    xlsxbook& book,
    unsigned long long int& i
    ) {
    parseAddress(cell, sheet, i);
    cacheValue  (cell, sheet, book, i); // Also caches format, as inextricable
    cacheFormula(cell, sheet, i);
}

// Based on hadley/readxl
// Get the A1-style address, and parse it for the row and column numbers.
// Simple parser: does not check that order of numbers and letters is correct
// row_ and column_ are one-based
void xlsxcell::parseAddress(
    rapidxml::xml_node<>* cell,
    xlsxsheet* sheet,
    unsigned long long int& i
    ) {
  rapidxml::xml_attribute<>* r = cell->first_attribute("r");
  if (r == NULL)
    stop("Invalid cell: lacks 'r' attribute");

  std::string address = r->value(); // we need this std::string in a moment
  sheet->address_[i] = address;

  // Iterate though the A1-style address string character by character
  int col = 0;
  int row = 0;
  for(std::string::const_iterator iter = address.begin();
      iter != address.end(); ++iter) {
    if (*iter >= '0' && *iter <= '9') { // If it's a number
      row = row * 10 + (*iter - '0'); // Then multiply existing row by 10 and add new number
    } else if (*iter >= 'A' && *iter <= 'Z') { // If it's a character
      col = 26 * col + (*iter - 'A' + 1); // Then do similarly with columns
    }
  }
  sheet->col_[i] = col;
  sheet->row_[i] = row;

  // Look up any comment using the address
  std::map<std::string, std::string>& comments = sheet->comments_;
  std::map<std::string, std::string>::iterator it = comments.find(address);
  if(it != comments.end()) {
     sheet->comment_[i] = it->second;
  }
}

void xlsxcell::cacheValue(
    rapidxml::xml_node<>* cell,
    xlsxsheet* sheet,
    xlsxbook& book,
    unsigned long long int& i
    ) {
  // 'v' for 'value' is either literal (numeric) or an index into a string table
  rapidxml::xml_node<>* v = cell->first_node("v");
  std::string vvalue;
  if (v != NULL) {
    vvalue = v->value();
    sheet->content_[i] = vvalue;
  } else {
    sheet->content_[i] = NA_STRING;
  }

  // 't' for 'type' defines the meaning of 'v' for value
  rapidxml::xml_attribute<>* t = cell->first_attribute("t");
  std::string tvalue;
  if (t != NULL) {
    tvalue = t->value();
    sheet->type_[i] = tvalue;
  } else {
    sheet->type_[i] = NA_STRING;
  }

  // 's' for 'style' indexes into data structures of formatting
  rapidxml::xml_attribute<>* s = cell->first_attribute("s");
  // Default the local format id to '1' if not present
  int svalue;
  if (s != NULL) {
    svalue = strtol(s->value(), NULL, 10);
  } else {
    svalue = 0;
  }
  sheet->local_format_id_[i] = svalue + 1;
  sheet->style_format_[i] = book.styles_.cellStyles_[book.styles_.cellXfs_[svalue].xfId_[0]];

  if (t != NULL && tvalue == "inlineStr") {
    sheet->data_type_[i] = "character";
    rapidxml::xml_node<>* is = cell->first_node("is");
    if (is != NULL) { // Get the inline string if it's really there
      std::string inlineString;
      parseString(is, inlineString); // value is modified in place
      SET_STRING_ELT(sheet->character_, i, Rf_mkCharCE(inlineString.c_str(), CE_UTF8));
    }
    return;
  } else if (v == NULL) {
    // Can't now be an inline string (tested above)
    sheet->data_type_[i] = "blank";
    return;
  } else if (t == NULL || tvalue == "n") {
    if (book.styles_.cellXfs_[svalue].applyNumberFormat_[0] == 1) {
      // local number format applies
      if (book.styles_.isDate_[book.styles_.cellXfs_[svalue].numFmtId_[0]]) {
        // local number format is a date format
        sheet->data_type_[i] = "date";
        sheet->date_[i] = (strtof(vvalue.c_str(), NULL) - book.dateOffset_) * 86400;
        return;
      } else {
        sheet->data_type_[i] = "numeric";
        sheet->numeric_[i] = strtof(vvalue.c_str(), NULL);
      }
    } else if (
          book.styles_.isDate_[
            book.styles_.cellStyleXfs_[
              book.styles_.cellXfs_[svalue].xfId_[0]
            ].numFmtId_[0]
          ]
        ) {
      // style number format is a date format
      sheet->data_type_[i] = "date";
      sheet->date_[i] = (strtof(vvalue.c_str(), NULL) - book.dateOffset_) * 86400;
      return;
    } else {
      sheet->data_type_[i] = "numeric";
      sheet->numeric_[i] = strtof(vvalue.c_str(), NULL);
    }
  } else if (tvalue == "s") {
    // the t attribute exists and its value is exactly "s", so v is an index
    // into the string table.
    sheet->data_type_[i] = "character";
    sheet->character_[i] = book.strings_[strtol(vvalue.c_str(), NULL, 10)];
    return;
  } else if (tvalue == "str") {
    // Formula, which could have evaluated to anything, so only a string is safe
    sheet->data_type_[i] = "character";
    sheet->character_[i] = vvalue;
    return;
  } else if (tvalue == "b"){
    sheet->data_type_[i] = "logical";
    sheet->logical_[i] = strtof(vvalue.c_str(), NULL);
    return;
  } else if (tvalue == "e") {
    sheet->data_type_[i] = "error";
    sheet->error_[i] = vvalue;
    return;
  } else if (tvalue == "d") {
    // Does excel use this date type? Regardless, don't have cross-platform
    // ISO8601 parser (yet) so need to return as text.
    sheet->data_type_[i] = "date (ISO8601)";
    return;
  } else {
    sheet->data_type_[i] = "unknown";
    return;
  }
}

void xlsxcell::cacheFormula(
    rapidxml::xml_node<>* cell,
    xlsxsheet* sheet,
    unsigned long long int& i
    ) {
  // TODO: Formulas are more complicated than this, because they're shared.
  // p.1629 'shared' and 'si' attributes
  // TODO: Array formulas use the ref attribute for their range, and t to
  // state that they're 'array'.
  rapidxml::xml_node<>* f = cell->first_node("f");
  if (f != NULL) {
    sheet->formula_[i] = f->value();
    rapidxml::xml_attribute<>* f_t = f->first_attribute("t");
    if (f_t != NULL) {
      sheet->formula_type_[i] = f_t->value();
    } else {
      sheet->formula_type_[i] = NA_STRING;
    }
    rapidxml::xml_attribute<>* ref = f->first_attribute("ref");
    if (ref != NULL) {
      sheet->formula_ref_[i] = ref->value();
    } else {
      sheet->formula_ref_[i] = NA_STRING;
    }
    rapidxml::xml_attribute<>* si = f->first_attribute("si");
    if (si != NULL) {
      sheet->formula_group_[i] = strtol(si->value(), NULL, 10);
    } else {
      sheet->formula_group_[i] = NA_INTEGER;
    }
  } else {
    sheet->formula_[i] = NA_STRING;
    sheet->formula_type_[i] = NA_STRING;
    sheet->formula_ref_[i] = NA_STRING;
    sheet->formula_group_[i] = NA_INTEGER;
  }
}
