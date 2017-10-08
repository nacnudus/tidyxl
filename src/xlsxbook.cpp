#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxbook.h"
#include "xlsxsheet.h"
#include "styles.h"
#include "utils.h"

using namespace Rcpp;

xlsxbook::xlsxbook(const std::string& path): path_(path), styles_(path_) {
  std::string book = zip_buffer(path_, "xl/workbook.xml");

  rapidxml::xml_document<> xml;
  xml.parse<0>(&book[0]);

  rapidxml::xml_node<>* workbook = xml.first_node("workbook");

  cacheDateOffset(workbook); // Must come before cacheSheets
  cacheStrings();
}

xlsxbook::xlsxbook(
    const std::string& path,
    CharacterVector& sheet_paths,
    CharacterVector& sheet_names,
    CharacterVector& comments_paths):
  path_(path),
  sheet_paths_(sheet_paths),
  sheet_names_(sheet_names),
  comments_paths_(comments_paths),
  styles_(path_) {
  std::string book = zip_buffer(path_, "xl/workbook.xml");

  rapidxml::xml_document<> xml;
  xml.parse<0>(&book[0]);

  rapidxml::xml_node<>* workbook = xml.first_node("workbook");

  cacheDateOffset(workbook); // Must come before cacheSheets
  cacheStrings();
  cacheSheetXml();
  createSheets();
  countCells();
  initializeColumns();
  cacheInformation();
}

// Based on hadley/readxl
void xlsxbook::cacheStrings() {
  if (!zip_has_file(path_, "xl/sharedStrings.xml"))
    return;

  std::string xml = zip_buffer(path_, "xl/sharedStrings.xml");
  rapidxml::xml_document<> sharedStrings;
  sharedStrings.parse<0>(&xml[0]);

  rapidxml::xml_node<>* sst = sharedStrings.first_node("sst");
  rapidxml::xml_attribute<>* uniqueCount = sst->first_attribute("uniqueCount");
  if (uniqueCount != NULL) {
    unsigned long int n = strtol(uniqueCount->value(), NULL, 10);
    strings_.reserve(n);
  }

  // 18.4.8 si (String Item) [p1725]
  for (rapidxml::xml_node<>* string = sst->first_node();
      string; string = string->next_sibling()) {
    std::string out;
    parseString(string, out);    // missing strings are treated as empty ""
    strings_.push_back(out);
  }
}

void xlsxbook::cacheDateOffset(rapidxml::xml_node<>* workbook) {
  rapidxml::xml_node<>* workbookPr = workbook->first_node("workbookPr");
  if (workbookPr != NULL) {
    rapidxml::xml_attribute<>* date1904 = workbookPr->first_attribute("date1904");
    if (date1904 != NULL) {
      std::string is1904 = date1904->value();
      if ((is1904 == "1") || (is1904 == "true")) {
        dateSystem_ = 1904;
        dateOffset_ = 24107;
        return;
      }
    }
  }

  dateSystem_ = 1900;
  dateOffset_ = 25569;
}

void xlsxbook::cacheSheetXml() {
  // Loop through sheets, reading the xml into memory
  CharacterVector::iterator in_it;
  for(in_it = sheet_paths_.begin(); in_it != sheet_paths_.end(); ++in_it) {
    std::string xml(*in_it);
    sheet_xml_.push_back(zip_buffer(path_, xml));
  }
}

void xlsxbook::createSheets() {
  // Loop through sheets
  std::vector<std::string>::iterator xml;
  CharacterVector::iterator name;
  CharacterVector::iterator comments_path;
  int i = 0;
  for(xml = sheet_xml_.begin(),
      name = sheet_names_.begin(),
      comments_path = comments_paths_.begin();
      xml != sheet_xml_.end();
      ++xml, ++name, ++comments_path) {
    std::string xmlstring(*xml);
    Rcpp::String namestring(*name);
    Rcpp::String comments_path_string(*comments_path);
    sheets_.emplace_back(xlsxsheet(namestring, xmlstring, *this, comments_path_string));
    ++i;
  }
}

void xlsxbook::countCells() {
  // Loop through sheets, adding their cellcounts
  cellcount_ = 0;
  std::vector<xlsxsheet>::iterator sheet;
  for(sheet = sheets_.begin();
      sheet != sheets_.end();
      ++sheet) {
    cellcount_ += sheet->cellcount_;
  }
}

void xlsxbook::initializeColumns() {
  sheet_           = CharacterVector(cellcount_, NA_STRING);
  address_         = CharacterVector(cellcount_, NA_STRING);
  row_             = IntegerVector(cellcount_,   NA_INTEGER);
  col_             = IntegerVector(cellcount_,   NA_INTEGER);
  is_blank_        = LogicalVector(cellcount_,   false);
  data_type_       = CharacterVector(cellcount_, NA_STRING);
  error_           = CharacterVector(cellcount_, NA_STRING);
  logical_         = LogicalVector(cellcount_,   NA_LOGICAL);
  numeric_         = NumericVector(cellcount_,   NA_REAL);
  date_            = NumericVector(cellcount_,   NA_REAL);
  date_.attr("class") = CharacterVector::create("POSIXct", "POSIXt");
  date_.attr("tzone") = "UTC";
  character_       = CharacterVector(cellcount_, NA_STRING);
  formula_         = CharacterVector(cellcount_, NA_STRING);
  formula_type_    = CharacterVector(cellcount_, NA_STRING);
  formula_ref_     = CharacterVector(cellcount_, NA_STRING);
  formula_group_   = IntegerVector(cellcount_,   NA_INTEGER);
  comment_         = CharacterVector(cellcount_, NA_STRING);
  height_          = NumericVector(cellcount_,   NA_REAL);
  width_           = NumericVector(cellcount_,   NA_REAL);
  style_format_    = CharacterVector(cellcount_, NA_STRING);
  local_format_id_ = IntegerVector(cellcount_,   NA_INTEGER);
}

void xlsxbook::cacheInformation() {
  // Loop through sheets
  List sheet_list(sheet_paths_.size());

  std::vector<std::string>::iterator xml;
  std::vector<xlsxsheet>::iterator sheet;
  List::iterator sheet_list_it;

  unsigned long long int i(0); // position of each cell in the output vectors
  for(xml = sheet_xml_.begin(),
      sheet = sheets_.begin(),
      sheet_list_it = sheet_list.begin();
      xml != sheet_xml_.end();
      ++xml, ++sheet, ++sheet_list_it) {
    rapidxml::xml_document<> doc;
    doc.parse<0>(&(*xml)[0]);
    rapidxml::xml_node<>* workbook = doc.first_node("worksheet");
    rapidxml::xml_node<>* sheetData = workbook->first_node("sheetData");
    sheet->parseSheetData(sheetData, i);
    sheet->appendComments(i);
  }

  // Returns a nested data frame of everything, the data frame itself wrapped in
  // a list.
  information_ = List::create(
      _["sheet"] = sheet_,
      _["address"] = address_,
      _["row"] = row_,
      _["col"] = col_,
      _["is_blank"] = is_blank_,
      _["data_type"] = data_type_,
      _["error"] = error_,
      _["logical"] = logical_,
      _["numeric"] = numeric_,
      _["date"] = date_,
      _["character"] = character_,
      _["formula"] = formula_,
      _["formula_type"] = formula_type_,
      _["formula_ref"] = formula_ref_,
      _["formula_group"] = formula_group_,
      _["comment"] = comment_,
      _["height"] = height_,
      _["width"] = width_,
      _["style_format"] = style_format_,
      _["local_format_id"] = local_format_id_);

  // Turn list of vectors into a data frame without checking anything
  int n = Rf_length(information_[0]);
  information_.attr("class") = Rcpp::CharacterVector::create("tbl_df", "tbl", "data.frame");
  information_.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -n); // Dunno how this works (the -n part)
}
