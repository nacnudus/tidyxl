#include <Rcpp.h>
#include "shared_formula.h"
#include "ref_grammar.h"
#include "ref.h"

shared_formula::shared_formula(
    std::string& text,
    int& row,
    int& col
    ): text_(text), row_(row), col_(col) {
  memory_input<> in_mem(text_, "original-formula");
  parse< xlref::root, xlref::tokenize >( in_mem, types_, tokens_, refs_);
}

std::string shared_formula::offset(int& row, int& col) const {
  std::string out;

  // Size of offset
  int rows = row - row_;
  int cols = col - col_;

  std::vector<std::string>::const_iterator i_token = tokens_.begin();
  std::vector<ref>::const_iterator i_ref = refs_.begin();

  for(std::vector<token_type>::const_iterator i_type = types_.begin();
      i_type != types_.end(); ++i_type) {
    switch (*i_type) {
      case token_type::REF:
        out = out + (*i_ref).offset(rows, cols);
        ++i_ref;
        break;

      default:
        out = out + *i_token;
        ++i_token;
        break;
    }
  }

  return out;
}

