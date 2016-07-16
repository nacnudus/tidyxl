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
    parseAddress(address_, row_, col_);

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
