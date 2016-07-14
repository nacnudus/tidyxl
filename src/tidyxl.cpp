#include <algorithm>
#include <Rcpp.h>
#include "xlsxsheet.h"

using namespace Rcpp;

// This is the entry-point from R into Rcpp.  The main function, xlsx_read_,
// instantiates the xlsxbook class, and then passes that (by reference?
// pointer?) to one or more xlsxsheet instances, iterating through them and
// returning a list of 'sheets'.

// Currently, only one sheet is handled at a time, but it is returned wrapped in
// a list, to prevent multiple sheets breaking the api.

// [[Rcpp::export]]
List xlsx_read_(std::string path, IntegerVector sheets) {
  // Parse book-level information (e.g. styles, themes, strings, date system)
  /* xlsxbook book(path); */

  // Loop through sheets
  List out(sheets.size());

  IntegerVector::iterator in_it;
  List::iterator out_it;

  for(in_it = sheets.begin(), out_it = out.begin(); in_it != sheets.end(); 
      ++in_it, ++out_it) {
    *out_it = xlsxsheet(path, *in_it).information();
  }

  return out;
}

