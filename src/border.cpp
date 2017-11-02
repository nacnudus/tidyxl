#include <Rcpp.h>
#include "rapidxml.h"
#include "border.h"
#include "xlsxstyles.h"
#include "stroke.h"

using namespace Rcpp;

inline int bool_value(rapidxml::xml_node<>* node, const char* name) {
  std::string value;
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute == NULL) {
    return(false);
  }
  value = attribute->value();
  if (value == "0" || value == "false") {
    return(false);
  }
  return(true);
}

border::border() {} // default constructor

border::border(rapidxml::xml_node<>* border,
    xlsxstyles* styles
    ):
  left_(border->first_node("left"), styles),
  right_(border->first_node("right"), styles),
  start_(border->first_node("start"), styles), // gnumeric
  end_(border->first_node("end"), styles),     // gnumeric
  top_(border->first_node("top"), styles),
  bottom_(border->first_node("bottom"), styles),
  diagonal_(border->first_node("diagonal"), styles),
  vertical_(border->first_node("vertical"), styles),    // can't produce in Excel
  horizontal_(border->first_node("horizontal"), styles) // can't produce in Excel
{
  diagonalDown_ = bool_value(border, "diagonalDown");
  diagonalUp_ = bool_value(border, "diagonalUp");

  // I haven't been able to create an outline attribute in Excel 2016
  outline_ = bool_value(border, "outline");
}
