#ifndef XLSXCELL_
#define XLSXCELL_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"

class xlsxcell {

rapidxml::xml_node<>* c_; // Provided to constructor
xlsxbook& book_;          // Parent workbook

rapidxml::xml_attribute<>* r_; // Address
rapidxml::xml_attribute<>* s_; // Style index
rapidxml::xml_attribute<>* t_; // Data type
rapidxml::xml_node<>* v_;      // Contents/value
rapidxml::xml_node<>* f_;      // Formula

const char* vvalue_; // value of v node, optimises cacheString
const char* tvalue_; // value of t node, optimises cacheString

/* celltype celltype;     // TODO: cell type enumeration */

// The remaining variables go to R.

std::string  address_;    // Value of node r
int          row_;        // Parsed address_ (one-based)
int          col_;        // Parsed address_ (one-based)

Rcpp::String content_;    // Unparsed value of node v
Rcpp::String formula_;    // If present
int          formula_group_; // Value of node "si", if present, indicates shared formulas
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
int style_format_id_;      // cellXfs xfId links to cellStyleXfs entry
int local_format_id_;      // cell 'c' links to cellXfs entry

public:

  xlsxcell(rapidxml::xml_node<>* c, double& height, 
    std::vector<double>& colWidths, xlsxbook& book); // c is the node of the cell
  std::string&  address();
  int&          row();
  int&          col();
  Rcpp::String& content();
  Rcpp::String& formula();
  int&          formula_group();
  Rcpp::String& type();
  Rcpp::String& character();
  double&       height();
  double&       width();
  int&          style_format_id();
  int&          local_format_id();

private:

  void cacheString();
  void cacheFormat(); // Only the index of the local and style formats (cellXfs and cellStyleXfs)

};

#endif

