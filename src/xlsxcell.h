#ifndef XLSXCELL_
#define XLSXCELL_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"

class xlsxcell {

/* celltype celltype;     // TODO: cell type enumeration */

// The remaining variables go to R.

std::string  address_;    // Value of node r
int          row_;        // Parsed address_ (one-based)
int          col_;        // Parsed address_ (one-based)

Rcpp::String content_;    // Unparsed value of node v
Rcpp::String formula_;    // If present
Rcpp::String formula_type_; // value of attribute "t", if present, indicates 'shared' or 'array'
Rcpp::String formula_ref_;  // value of attribute "ref", if present, indicates range of shared/array formula
int          formula_group_; // Value of attribute "si", if present, indicates shared formulas
Rcpp::String type_;       // Type of the parsed value, not exactly node t though

Rcpp::List   value_;      // Parsed value wrapped in unnamed list

// The following are NA unless relevant to the cell type, but if NA then
// they should never be sought, because celltype should flag which is relevant.
bool         logical_;    // Parsed value
double       numeric_;    // Parsed value
double       date_;       // Parsed value
Rcpp::String character_;  // Parsed value
Rcpp::String error_;      // Parsed value

// The following are always used.
double height_;           // Provided to constructor
double width_;            // Provided to constructor
unsigned long int style_format_id_;      // cellXfs xfId links to cellStyleXfs entry
unsigned long int local_format_id_;      // cell 'c' links to cellXfs entry

public:

  xlsxcell(rapidxml::xml_node<>* cell, double& height, 
    std::vector<double>& colWidths, xlsxbook& book); // 'cell' is the cell node,
                                                     // 'book' is the parent workbook
  std::string&  address();
  int&          row();
  int&          col();
  Rcpp::String& content();
  Rcpp::String& formula();
  Rcpp::String& formula_type();
  Rcpp::String& formula_ref();
  int&          formula_group();
  Rcpp::String& type();
  Rcpp::String& character();
  double&       height();
  double&       width();
  unsigned long int& style_format_id();
  unsigned long int& local_format_id();

private:

  void cacheString(
      rapidxml::xml_node<>* cell,
      xlsxbook& book, 
      rapidxml::xml_node<>* v,
      rapidxml::xml_attribute<>* t
      ); // Local and style format indexes
  void cacheFormat(rapidxml::xml_node<>* cell, xlsxbook& book); // Only the index of the local and style formats (cellXfs and cellStyleXfs)
  void cacheFormula(rapidxml::xml_node<>* cell);

};

#endif

