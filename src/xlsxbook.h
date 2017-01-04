#ifndef XLSXBOOK_
#define XLSXBOOK_

#include <Rcpp.h>
#include "rapidxml.h"
#include "styles.h"

class xlsxbook {

  public:

    const std::string& path_;                     // workbook path
    std::vector<std::string> sheets_;             // sheet names
    std::vector<std::string> strings_;            // strings table

    styles styles_;

    int dateOffset_; // for converting 1900 or 1904 Excel datetimes to R

    xlsxbook(const std::string& path);    // constructor


    void cacheSheets(rapidxml::xml_node<>* sheets); // sheet names, not data
    void cacheStrings();
    void cacheStyles();
    void cacheDateOffset(rapidxml::xml_node<>* workbook);

};

#endif
