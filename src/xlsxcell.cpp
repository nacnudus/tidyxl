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
    cacheValue  (cell, sheet, book, i);
    cacheFormat (cell, sheet, book, i); // local and style format indexes
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
}

// Based on hadley/readxl
// Get the value of the cell, which could be numeric, a string, or an index into
// the string table.
void xlsxcell::cacheValue(
    rapidxml::xml_node<>* cell,
    xlsxsheet* sheet,
    xlsxbook& book,
    unsigned long long int& i
    ) {
  // 'v' for 'value' is either literal (numeric) or an index into a string table
  rapidxml::xml_node<>* v = cell->first_node("v");
  if (v != NULL) {
    sheet->content_[i] = v->value();
  } else {
    sheet->content_[i] = NA_STRING;
  }

  // 't' for 'type' defines the meaning of 'v' for value
  rapidxml::xml_attribute<>* t = cell->first_attribute("t");
  if (t != NULL) {
    sheet->type_[i] = t->value();
  } else {
    sheet->type_[i] = NA_STRING;
  }

  // If an inline string, then it must be parsed, if a string in the string 
  // table, then that string must be obtained.

  // Try the string table
  if (v != NULL && t != NULL && strncmp(t->value(), "s", t->value_size()) == 0) {
    // the t attribute exists and its value is exactly "s", so v is an index
    // into the string table.
    sheet->character_[i] = book.strings_[strtol(v->value(), NULL, 10)];
    return;
  }

  // Otherwise, check whether it's an inline string instead
  // We could check for t="inlineString" or the presence of child "is".  We do
  // the latter, same as hadley/readxl.
  // Is it an inline string?  // 18.3.1.53 is (Rich Text Inline) [p1649]
  rapidxml::xml_node<>* is = cell->first_node("is");
  if (is != NULL) {
    std::string inlineString;
    if (!parseString(is, inlineString)) { // value is modified in place
      sheet->character_[i] = NA_STRING;
    } else {
      sheet->character_[i] = inlineString;
    }
    return;
  }

  // Otherwise, it's neither in the string table nor an inline string, so it 
  // isn't a string at all.
  sheet->character_[i] = NA_STRING;
}

void xlsxcell::cacheFormat(
    rapidxml::xml_node<>* cell,
    xlsxsheet* sheet,
    xlsxbook& book,
    unsigned long long int& i
    ) {
  rapidxml::xml_attribute<>* s = cell->first_attribute("s");
  // Default the local format id to '1' if not present
  int local_format_id;
  if (s != NULL) {
    local_format_id = strtol(s->value(), NULL, 10) + 1;
  } else {
    local_format_id = 1;
  }
  sheet->local_format_id_[i] = local_format_id;
  sheet->style_format_id_[i] = book.cellXfs_xfId_[local_format_id - 1];
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
