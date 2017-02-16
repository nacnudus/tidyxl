#ifndef XLSXBOOK_
#define XLSXBOOK_

#include <Rcpp.h>
#include "rapidxml.h"
#include "styles.h"

class xlsxbook {

  public:

    const std::string& path_;                     // workbook path
    std::vector<std::string> strings_;            // strings table

    styles styles_;

    int dateSystem_; // 1900 or 1904
    int dateOffset_; // for converting 1900 or 1904 Excel datetimes to R

    xlsxbook(const std::string& path);    // constructor

    void cacheStrings();
    void cacheStyles();
    void cacheDateOffset(rapidxml::xml_node<>* workbook);

};

#endif
