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

// Currently, only one sheet is handled at a time, but it is returned wrapped in
// a list, to prevent multiple sheets breaking the api.

// [[Rcpp::export]]
CharacterVector xlsx_sheet_names(std::string path) {
  // Return the names of the sheets in a book.
  // Duplicates xlsxbook.sheets() to avoid parsing anything but the sheets.

  std::string book = zip_buffer(path, "xl/workbook.xml");

  rapidxml::xml_document<> bookXml;
  bookXml.parse<0>(&book[0]);

  rapidxml::xml_node<>* rootNode = bookXml.first_node("workbook");
  if (rootNode == NULL)
    Rcpp::stop("Invalid workbook xml (no <workbook>)");

  rapidxml::xml_node<>* sheets = rootNode->first_node("sheets");
  if (sheets == NULL)
    Rcpp::stop("Invalid workbook xml (no <sheets>)");

  std::vector<std::string> names;

  for (rapidxml::xml_node<>* sheet = sheets->first_node();
      sheet; sheet = sheet->next_sibling()) {
    std::string name = sheet->first_attribute("name")->value();
    names.push_back(name);
  }

  return wrap(names);
}

// [[Rcpp::export]]
List xlsx_read_(std::string path, IntegerVector sheets) {
  // Parse book-level information (e.g. styles, themes, strings, date system)
  xlsxbook book(path);

  List out;
  out["sheets"] = xlsxbook(path).sheets();

  /* // Loop through sheets */
  /* List out(sheets.size()); */

  /* IntegerVector::iterator in_it; */
  /* List::iterator out_it; */

  /* for(in_it = sheets.begin(), out_it = out.begin(); in_it != sheets.end(); */ 
  /*     ++in_it, ++out_it) { */
  /*   *out_it = xlsxsheet(path, *in_it).information(); */
  /* } */

  return out;
}
