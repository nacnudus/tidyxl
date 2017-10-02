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

    Rcpp::List information_; // dataframes of sheet contents

    xlsxbook(const std::string& path);    // constructor

    xlsxbook(
        const std::string& path,
        Rcpp::CharacterVector& sheet_names,
        Rcpp::CharacterVector& sheet_paths,
        Rcpp::CharacterVector& comments_paths
        );

    void cacheStrings();
    void cacheStyles();
    void cacheDateOffset(rapidxml::xml_node<>* workbook);
    void cacheInformation();

};

#endif
