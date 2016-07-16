#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxcell.h"
#include "utils.h"

using namespace Rcpp;

xlsxcell::xlsxcell(rapidxml::xml_node<>* c): c_(c) {
    rapidxml::xml_attribute<>* r = c_->first_attribute("r");
    if (r == NULL)
      Rcpp::stop("Invalid cell: lacks 'r' attribute");
    address_ = std::string(r->value());
    parseAddress();

    rapidxml::xml_node<>* v = c_->first_node("v");
    if (v == NULL) {
      content_ = NA_STRING;
    } else {
      content_ = v->value();
    }

    /* rapidxml::xml_attribute<>* type = cell_->first_attribute("t"); */
    /* if (type == NULL) { */
    /*   has_type_ = false; */
    /* } else { */
    /*   has_type_ = true; */
    /*   type_ = std::string(type->value()); */
    /* } */

    /* rapidxml::xml_node<>* formula = cell_->first_node("f"); */
    /* if (formula == NULL) { */
    /*   has_formula_ = false; */
    /* } else { */
    /*   has_formula_ = true; */
    /*   formula_ = std::string(formula->value()); */
    /* } */

    /* rapidxml::xml_attribute<>* style = cell_->first_attribute("s"); */
    /* if (style == NULL) {style_ = 0;} else {style_ = atoi(style->value());} */

}

std::string xlsxcell::address() {
  return(address_);
}

int xlsxcell::row() {
  return(row_);
}

int xlsxcell::col() {
  return(col_);
}

String xlsxcell::content() {
  return content_;
}

// Based on hadley/readxl
// Simple parser: does not check that order of numbers and letters is correct
// row_ and column_ are one-based
void xlsxcell::parseAddress() {
  col_ = 0;
  row_ = 0;

  // Iterate though std::string character by character
  for(std::string::const_iterator iter = address_.begin();
      iter != address_.end(); ++iter) {
    if (*iter >= '0' && *iter <= '9') { // If it's a number
      row_ = row_ * 10 + (*iter - '0'); // Then multiply existing row by 10 and add new number
    } else if (*iter >= 'A' && *iter <= 'Z') { // If it's a character
      col_ = 26 * col_ + (*iter - 'A' + 1); // Then do similarly  with columns
    } else {
      Rcpp::stop("Invalid character '%s' in cell ref '%s'", *iter, address_);
    }
  }
}

