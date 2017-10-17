#ifndef XLSXBOOK_
#define XLSXBOOK_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxsheet.h"
#include "xlsxstyles.h"

class xlsxbook {

  public:

    const std::string& path_;              // workbook path
    Rcpp::CharacterVector sheet_paths_;    // worksheet paths
    Rcpp::CharacterVector sheet_names_;    // worksheet names
    Rcpp::CharacterVector comments_paths_; // comments files
    std::vector<std::string> strings_;     // strings table

    xlsxstyles styles_;

    int dateSystem_; // 1900 or 1904
    int dateOffset_; // for converting 1900 or 1904 Excel datetimes to R

    std::vector<std::string> sheet_xml_; // xml of worksheets
    std::vector<xlsxsheet> sheets_;      // worksheet objects
    unsigned long long int cellcount_;   // total cellcount of all sheets

    Rcpp::List information_;             // dataframes of cells

    // Vectors of cell properties, to be wrapped in a data frame
    Rcpp::CharacterVector sheet_;     // The worksheet that a cell is in
    Rcpp::CharacterVector address_;   // Value of cell node r
    Rcpp::IntegerVector   row_;       // Parsed address_ (one-based)
    Rcpp::IntegerVector   col_;       // Parsed address_ (one-based)
    Rcpp::LogicalVector   is_blank_;
    Rcpp::CharacterVector data_type_; // Type of the parsed value
    Rcpp::CharacterVector error_;     // Parsed value
    Rcpp::LogicalVector   logical_;   // Parsed value
    Rcpp::NumericVector   numeric_;   // Parsed value
    Rcpp::NumericVector   date_;      // Parsed value
    Rcpp::CharacterVector character_; // Parsed value
    Rcpp::CharacterVector formula_;   // If present
    Rcpp::LogicalVector   is_array_;  // If formulaType is present
    Rcpp::CharacterVector formula_ref_;   // If present
    Rcpp::IntegerVector   formula_group_; // If present
    Rcpp::CharacterVector comment_;
    Rcpp::NumericVector   height_;          // Provided to cell constructor
    Rcpp::NumericVector   width_;           // Provided to cell constructor
    Rcpp::CharacterVector style_format_;    // cellXfs xfId links to cellStyleXfs entry
    Rcpp::IntegerVector   local_format_id_; // cell 'c' links to cellXfs entry

    xlsxbook(const std::string& path);    // constructor

    xlsxbook(
        const std::string& path,
        Rcpp::CharacterVector& sheet_names,
        Rcpp::CharacterVector& sheet_paths,
        Rcpp::CharacterVector& comments_paths
        );

    void cacheStrings();
    void cacheDateOffset(rapidxml::xml_node<>* workbook);
    void cacheSheetXml();
    void createSheets();
    void countCells();
    void initializeColumns();
    void cacheCells();
    void cacheInformation();

};

#endif
