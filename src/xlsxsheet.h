#ifndef XLSXSHEET_
#define XLSXSHEET_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"
#include "shared_formula.h"

class xlsxbook;

class xlsxsheet {

  public:

    std::string name_;

    unsigned long long int i_; // cell counter
    double defaultRowHeight_;
    double defaultColWidth_;
    std::vector<double> colWidths_;
    std::vector<double> rowHeights_;
    std::map<int, shared_formula> shared_formulas_;
    xlsxbook& book_; // reference to parent workbook
    Rcpp::List information_; // Wrapper for variables returned to R

    // The remaining variables go to R
    Rcpp::CharacterVector address_;   // Value of cell node r
    Rcpp::IntegerVector   row_;       // Parsed address_ (one-based)
    Rcpp::IntegerVector   col_;       // Parsed address_ (one-based)
    Rcpp::CharacterVector content_;   // Unparsed value of cell node v
    Rcpp::CharacterVector formula_;   // If present
    Rcpp::CharacterVector formula_type_; // If present
    Rcpp::CharacterVector formula_ref_;  // If present
    Rcpp::IntegerVector   formula_group_; // If present

    Rcpp::List  value_;               // Parsed values wrapped in unnamed lists
    Rcpp::CharacterVector type_;      // value of cell node t
    Rcpp::CharacterVector data_type_; // Type of the parsed value

    Rcpp::CharacterVector error_;     // Parsed value
    Rcpp::LogicalVector   logical_;   // Parsed value
    Rcpp::NumericVector   numeric_;   // Parsed value
    Rcpp::NumericVector   date_;      // Parsed value
    Rcpp::CharacterVector character_; // Parsed value

    std::map<std::string, std::string> comments_;
    Rcpp::CharacterVector comment_;

    // The following are always used.
    Rcpp::NumericVector   height_;          // Provided to cell constructor
    Rcpp::NumericVector   width_;           // Provided to cell constructor
    Rcpp::CharacterVector style_format_;    // cellXfs xfId links to cellStyleXfs entry
    Rcpp::IntegerVector   local_format_id_; // cell 'c' links to cellXfs entry

    xlsxsheet(
        const std::string& name,
        std::string& sheet_xml,
        xlsxbook& book,
        Rcpp::String comments_path);
    Rcpp::List& information();       // Cells contents and styles DF wrapped in list

    void cacheDefaultRowColDims(rapidxml::xml_node<>* worksheet);
    void cacheColWidths(rapidxml::xml_node<>* worksheet);
    unsigned long long int cacheCellcount(rapidxml::xml_node<>* sheetData);
    void cacheComments(Rcpp::String comments_path);
    void initializeColumns(rapidxml::xml_node<>* sheetData);
    void parseSheetData(rapidxml::xml_node<>* sheetData);
    void appendComments();

};

#endif
