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
    int defaultColOutlineLevel_;
    int defaultRowOutlineLevel_;
    // TODO: Should the below vectors should be unsigned?
    std::vector<double> colWidths_;
    std::vector<double> rowHeights_;
    std::vector<int> colOutlineLevels_;
    std::vector<int> rowOutlineLevels_;
    std::map<int, shared_formula> shared_formulas_;
    xlsxbook& book_; // reference to parent workbook
    std::map<std::string, std::string> comments_; // lookup table of comments
    bool include_blank_cells_; // whether to include cells with no value

    xlsxsheet(
        const std::string& name,
        std::string& sheet_xml,
        xlsxbook& book,
        Rcpp::String comments_path,
        const bool& include_blank_cells);

    void cacheDefaultRowColAttributes(rapidxml::xml_node<>* worksheet);
    void cacheColAttributes(rapidxml::xml_node<>* worksheet);
    unsigned long long int cacheCellcount(rapidxml::xml_node<>* sheetData);
    void cacheComments(Rcpp::String comments_path);
    void parseSheetData(
        rapidxml::xml_node<>* sheetData,
        unsigned long long int& i);
    void appendComments(unsigned long long int& i);

};

#endif
