#ifndef SHARED_FORMULA_
#define SHARED_FORMULA_

#include <Rcpp.h>
#include "ref_grammar.h"
#include "ref.h"

class shared_formula {

  std::string text_;

  // It's position, from which dependent formulas are offset
  int row_;
  int col_;

  std::vector<token_type> types_;
  std::vector<std::string> tokens_;
  std::vector<ref> refs_;

  public:

    shared_formula(std::string& text, int& row, int& col);

    std::string offset(int& row, int& col) const; // return offsetted address
};

#endif
