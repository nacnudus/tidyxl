#include <Rcpp.h>
#include "rapidxml.h"
#include "fill.h"
#include "styles.h"
#include "patternFill.h"
#include "gradientFill.h"

using namespace Rcpp;

fill::fill() {} // Default constructor

fill::fill(rapidxml::xml_node<>* fill,
    styles* styles
    ):
  patternFill_(fill->first_node("patternFill"), styles),
  gradientFill_(fill->first_node("gradientFill"), styles)
{}
