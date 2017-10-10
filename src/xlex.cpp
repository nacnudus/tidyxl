#include <Rcpp.h>
#include "token_grammar.h"
#include "paren_type.h"

// [[Rcpp::export]]
Rcpp::List xlex_(Rcpp::CharacterVector x)
{

  std::string in_string;                   // the formula from the input x
  Rcpp::List out;                          // wraps types, tokens, levels
  std::vector<std::string> types;          // bool, number, text, function, etc.
  std::vector<std::string> tokens;         // the tokens themselves
  std::vector<int> levels;                 // level within nested formula
  int level(0);                            // starting level
  std::vector<paren_type> fun_paren;       // what is the context of a comma?

  in_string = Rcpp::as<std::string>(x);

  // default context before any '('.  Never required in valid formulas, but
  // avoids hard crash in the event, e.g. a formula =A1,B2.
  fun_paren.push_back(paren_type::PARENTHESES);

  memory_input<> in_mem(in_string, "original-formula");
  parse< xltoken::root, xltoken::tokenize >(in_mem,
                                            level,
                                            levels,
                                            fun_paren,
                                            types,
                                            tokens);

  out = Rcpp::List::create(
      Rcpp::_["level"] = levels,
      Rcpp::_["type"] = types,
      Rcpp::_["token"] = tokens
      );

  int n = tokens.size();
  out.attr("class") = Rcpp::CharacterVector::create("xlex", "tbl_df", "tbl", "data.frame");
  out.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -n); // Dunno how this works (the -n part)

  return out;
}


