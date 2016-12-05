#ifndef XLSXBOOK_
#define XLSXBOOK_

#include <Rcpp.h>
#include "rapidxml.h"

class xlsxbook {

const std::string& path_;
std::vector<std::string> sheets_;
std::vector<std::string> strings_;
std::vector<std::string>::size_type strings_size_;
std::vector<int> cellXfs_xfId_;

public:

  xlsxbook(const std::string& path);    // constructor
  std::string path();                   // workbook path
  std::vector<std::string>& sheets();   // sheet names
  std::vector<std::string>& strings();  // strings table
  std::vector<std::string>::size_type strings_size(); // length of strings table
  std::vector<int>& cellXfs_xfId();     // link from cellXfs from cellStyleXfs

private:

  void cacheSheets(rapidxml::xml_node<>* sheets); // sheet names
  void cacheStrings();
  void cacheCellXfsXfId(); // 'xfId' of 'cellXfs' in 'styles.xml' to link cells with themes

};

#endif
