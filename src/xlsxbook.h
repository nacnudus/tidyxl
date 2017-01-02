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

    xlsxbook(const std::string& path);    // constructor


    void cacheSheets(rapidxml::xml_node<>* sheets); // sheet names, not data
    void cacheStrings();
    void cacheStyles();

};

#endif
