#ifndef XLSXBOOK_
#define XLSXBOOK_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxsheet.h"
#include "styles.h"

class xlsxbook {

  public:

    const std::string& path_;                     // workbook path
    Rcpp::CharacterVector sheet_paths_;                 // worksheet paths
    Rcpp::CharacterVector sheet_names_;                 // worksheet names
    Rcpp::CharacterVector comments_paths_;              // comments files
    std::vector<std::string> strings_;            // strings table

    styles styles_;

    int dateSystem_; // 1900 or 1904
    int dateOffset_; // for converting 1900 or 1904 Excel datetimes to R

    std::vector<std::string> sheet_xml_; // xml of worksheets
    std::vector<xlsxsheet> sheets_;      // worksheet objects
    unsigned long long int cellcount_;   // total cellcount of all sheets

    Rcpp::List information_; // dataframes of sheet contents

    // Vectors of cell properties, to be wrapped in a data frame
    Rcpp::CharacterVector address_;   // Value of cell node r
    Rcpp::IntegerVector   row_;       // Parsed address_ (one-based)
    Rcpp::IntegerVector   col_;       // Parsed address_ (one-based)
    Rcpp::CharacterVector content_;   // Unparsed value of cell node v
    Rcpp::CharacterVector formula_;   // If present
    Rcpp::CharacterVector formula_type_;  // If present
    Rcpp::CharacterVector formula_ref_;   // If present
    Rcpp::IntegerVector   formula_group_; // If present
    Rcpp::List            value_;     // Parsed values wrapped in unnamed lists
    Rcpp::CharacterVector type_;      // value of cell node t
    Rcpp::CharacterVector data_type_; // Type of the parsed value
    Rcpp::CharacterVector error_;     // Parsed value
    Rcpp::LogicalVector   logical_;   // Parsed value
    Rcpp::NumericVector   numeric_;   // Parsed value
    Rcpp::NumericVector   date_;      // Parsed value
    Rcpp::CharacterVector character_; // Parsed value
    std::map<std::string, std::string> comments_;
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
    void cacheInformation();

};

#endif
