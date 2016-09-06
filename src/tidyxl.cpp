#include <algorithm>
#include <Rcpp.h>
#include "zip.h"
#include "xlsxsheet.h"
#include "xlsxbook.h"

using namespace Rcpp;

// This is the entry-point from R into Rcpp.  The main function, xlsx_read_,
// instantiates the xlsxbook class, and then passes that (by reference?
// pointer?) to one or more xlsxsheet instances, iterating through them and
// returning a list of 'sheets'.

// [[Rcpp::export]]
List xlsx_read_(std::string path, IntegerVector sheets, CharacterVector names) {
  // Parse book-level information (e.g. styles, themes, strings, date system)
  xlsxbook book(path);

  // Loop through sheets
  List out(sheets.size());

  IntegerVector::iterator in_it;
  List::iterator out_it;

  for(in_it = sheets.begin(), out_it = out.begin(); in_it != sheets.end(); 
      ++in_it, ++out_it) {
    *out_it = xlsxsheet(*in_it, book).information();
  }

  out.attr("names") = names;

  return out;
}
