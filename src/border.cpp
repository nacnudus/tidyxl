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
  diagonalDown_ = IntegerVector::create(NA_INTEGER);
  rapidxml::xml_attribute<>* diagonalDown = border->first_attribute("diagonalDown");
  if (diagonalDown != NULL)
    diagonalDown_[0] = strtol(diagonalDown->value(), NULL, 10);

  diagonalUp_   = IntegerVector::create(NA_INTEGER);
  rapidxml::xml_attribute<>* diagonalUp = border->first_attribute("diagonalUp");
  if (diagonalUp != NULL)
    diagonalUp_[0] = strtol(diagonalUp->value(), NULL, 10);

  // I haven't been able to create an outline attribute in Excel 2016
  outline_      = IntegerVector::create(NA_INTEGER);
  rapidxml::xml_attribute<>* outline = border->first_attribute("outline");
  if (outline != NULL)
    outline_[0] = strtol(outline->value(), NULL, 10);
}
