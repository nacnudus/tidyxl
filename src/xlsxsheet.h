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

    unsigned long long int cellcount_; // allowing for comments on blank cells
    double defaultRowHeight_;
    double defaultColWidth_;
    std::vector<double> colWidths_;
    std::vector<double> rowHeights_;
    std::map<int, shared_formula> shared_formulas_;
    xlsxbook& book_; // reference to parent workbook
    std::map<std::string, std::string> comments_; // lookup table of comments

    xlsxsheet(
        const std::string& name,
        std::string& sheet_xml,
        xlsxbook& book,
        Rcpp::String comments_path);

    void cacheDefaultRowColDims(rapidxml::xml_node<>* worksheet);
    void cacheColWidths(rapidxml::xml_node<>* worksheet);
    unsigned long long int cacheCellcount(rapidxml::xml_node<>* sheetData);
    void cacheComments(Rcpp::String comments_path);
    void parseSheetData(
        rapidxml::xml_node<>* sheetData,
        unsigned long long int& i);
    void appendComments(unsigned long long int& i);

};

#endif
