#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"
#include "xlsxcell.h"
#include "utils.h"

using namespace Rcpp;

xlsxcell::xlsxcell(rapidxml::xml_node<>* cell,
    double& height, std::vector<double>& colWidths, xlsxbook& book): 
  height_(height) {
    r_ = cell->first_attribute("r");
    if (r_ == NULL)
      stop("Invalid cell: lacks 'r' attribute");
    address_ = std::string(r_->value());
    parseAddress(address_, row_, col_);
    width_ = colWidths[col_ - 1]; // row height is provided to the constructor
                                  // whereas col widths have to be looked up

    rapidxml::xml_node<>* v = cell->first_node("v");
    if (v != NULL) {
      content_ = v->value();    // holds the value of 'v' that will be returned
    } else {
      content_ = NA_STRING;
    }

    rapidxml::xml_attribute<>* t = cell->first_attribute("t");
    if (t != NULL) {
      type_ = t->value();       // holds the value of 't' that will be returned
    } else {
      type_ = NA_STRING;
    }

    // TODO: Formulas are more complicated than this, because they're shared.
    // p.1629 'shared' and 'si' attributes
    // TODO: Array formulas use the ref attribute for their range, and t to
    // state that they're 'array'.
    f_ = cell->first_node("f");
    if (f_ != NULL) {
      formula_ = f_->value();
      rapidxml::xml_attribute<>* t = f_->first_attribute("t");
      if (t != NULL) {
        formula_type_ = t->value();
      } else {
        formula_type_ = NA_STRING;
      }
      rapidxml::xml_attribute<>* ref = f_->first_attribute("ref");
      if (ref != NULL) {
        formula_ref_ = ref->value();
      } else {
        formula_ref_ = NA_STRING;
      }
      rapidxml::xml_attribute<>* si = f_->first_attribute("si");
      if (si != NULL) {
        formula_group_ = strtol(si->value(), NULL, 10);
      } else {
        formula_group_ = NA_INTEGER;
      }
    } else {
      formula_ = NA_STRING;
      formula_type_ = NA_STRING;
      formula_ref_ = NA_STRING;
      formula_group_ = NA_INTEGER;
    }

    cacheString(cell, book, v, t); // t_ and v_ must be obtained first
    cacheFormat(cell, book); // local and style format indexes
}

std::string& xlsxcell::address()         {return address_;}
int&         xlsxcell::row()             {return row_;}
int&         xlsxcell::col()             {return col_;}
String&      xlsxcell::content()         {return content_;}
String&      xlsxcell::formula()         {return formula_;}
String&      xlsxcell::formula_type()    {return formula_type_;}
String&      xlsxcell::formula_ref()     {return formula_ref_;}
int&         xlsxcell::formula_group()   {return formula_group_;}
String&      xlsxcell::type()            {return type_;}
String&      xlsxcell::character()       {return character_;}
double&      xlsxcell::height()          {return height_;}
double&      xlsxcell::width()           {return width_;}
unsigned long int& xlsxcell::style_format_id() {return style_format_id_;}
unsigned long int& xlsxcell::local_format_id() {return local_format_id_;}

// Based on hadley/readxl
void xlsxcell::cacheString(
    rapidxml::xml_node<>* cell,
    xlsxbook& book,
    rapidxml::xml_node<>* v,
    rapidxml::xml_attribute<>* t) {
  // If an inline string, it must be parsed, if a string in the string table, it
  // must be obtained.

  // We could check for t="inlineString" or the presence of child "is".  We do
  // the latter, same as hadley/readxl.
  // Is it an inline string?  // 18.3.1.53 is (Rich Text Inline) [p1649]
  rapidxml::xml_node<>* is = cell->first_node("is");
  if (is != NULL) {
    std::string inlineString;
    if (!parseString(is, inlineString)) { // value is modified in place
      character_ = NA_STRING;
    } else {
      character_ = inlineString;
    }
    return;
  }

  if (v != NULL && t != NULL && strncmp(tvalue_, "s", t->value_size()) == 0) {
    // the t attribute exists and its value is exactly "s", so v is an index
    // into the strings table.
    unsigned long int vvalue = strtol(v->value(), NULL, 10);
    const std::vector<std::string>& strings = book.strings();
    character_ = strings[vvalue];
    return;
  }
  character_ = NA_STRING;
}

void xlsxcell::cacheFormat(rapidxml::xml_node<>* cell, xlsxbook& book) {
    s_ = cell->first_attribute("s");
    local_format_id_ = (s_ != NULL) ? strtol(s_->value(), NULL, 10) + 1 : 1;
    style_format_id_ = book.cellXfs_xfId()[local_format_id_ - 1];
  }
