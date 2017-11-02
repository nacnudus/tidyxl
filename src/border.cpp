#include <Rcpp.h>
#include "rapidxml.h"
#include "border.h"
#include "xlsxstyles.h"
#include "stroke.h"

using namespace Rcpp;

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
  vertical_(border->first_node("vertical"), styles),
  horizontal_(border->first_node("horizontal"), styles)
{
  rapidxml::xml_attribute<>* diagonalDown = border->first_attribute("diagonalDown");
  if (diagonalDown != NULL) {
    diagonalDown_ = strtol(diagonalDown->value(), NULL, 10);
  } else {
    diagonalDown_ = NA_INTEGER;
  }

  rapidxml::xml_attribute<>* diagonalUp = border->first_attribute("diagonalUp");
  if (diagonalUp != NULL) {
    diagonalUp_ = strtol(diagonalUp->value(), NULL, 10);
  } else {
    diagonalUp_ = NA_INTEGER;
  }

  // I haven't been able to create an outline attribute in Excel 2016
  rapidxml::xml_attribute<>* outline = border->first_attribute("outline");
  if (outline != NULL) {
    outline_ = strtol(outline->value(), NULL, 10); // # nocov
  } else {
    outline_ = NA_INTEGER;
  }
}
