#include <algorithm>
#include <Rcpp.h>
#include "zip.h"
#include "xlsxsheet.h"
#include "xlsxbook.h"

using namespace Rcpp;

// This is the entry-point from R into Rcpp.  The main function, xlsx_read_,
// instantiates the xlsxbook class, and then passes that by reference to one or
// more xlsxsheet instances, iterating through them and returning a list of
// 'sheets'.

// [[Rcpp::export]]
List xlsx_read_(std::string path, IntegerVector sheets, CharacterVector names) {
  // Parse book-level information (e.g. styles, themes, strings, date system)
  xlsxbook book(path);

  // Loop through sheets
  List sheet_list(sheets.size());

  IntegerVector::iterator in_it;
  List::iterator sheet_list_it;

  for(in_it = sheets.begin(), sheet_list_it = sheet_list.begin(); in_it != sheets.end();
      ++in_it, ++sheet_list_it) {
    *sheet_list_it = xlsxsheet(*in_it, book).information();
  }

  sheet_list.attr("names") = names;

  List out = List::create(
      _["data"] = sheet_list,
      _["formats"] = List::create(
        _["local"] = book.styles_.local_,
        _["style"] = book.styles_.style_));

  return out;
}
