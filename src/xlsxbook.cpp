#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxbook.h"
#include "xlsxsheet.h"
#include "xlsxstyles.h"
#include "string.h"

using namespace Rcpp;

xlsxbook::xlsxbook(const std::string& path): path_(path), styles_(path_) {
  std::string book = zip_buffer(path_, "xl/workbook.xml");

  rapidxml::xml_document<> xml;
  xml.parse<rapidxml::parse_strip_xml_namespaces>(&book[0]);

  rapidxml::xml_node<>* workbook = xml.first_node("workbook");

  cacheDateOffset(workbook); // Must come before cacheSheets
  cacheStrings();
}

xlsxbook::xlsxbook(
    const std::string& path,
    CharacterVector& sheet_paths,
    CharacterVector& sheet_names,
    CharacterVector& comments_paths,
    const bool& include_blank_cells):
  path_(path),
  sheet_paths_(sheet_paths),
  sheet_names_(sheet_names),
  comments_paths_(comments_paths),
  styles_(path_),
  include_blank_cells_(include_blank_cells) {
  std::string book = zip_buffer(path_, "xl/workbook.xml");

  rapidxml::xml_document<> xml;
  xml.parse<rapidxml::parse_strip_xml_namespaces>(&book[0]);

  rapidxml::xml_node<>* workbook = xml.first_node("workbook");

  cacheDateOffset(workbook); // Must come before cacheSheets
  cacheStrings();
  cacheSheetXml();
  createSheets();
  countCells();
  initializeColumns();
  cacheInformation();
}

// Based on tidyverse/readxl
void xlsxbook::cacheStrings() {
  if (!zip_has_file(path_, "xl/sharedStrings.xml"))
    return;

  std::string xml = zip_buffer(path_, "xl/sharedStrings.xml");
  rapidxml::xml_document<> sharedStrings;
  sharedStrings.parse<rapidxml::parse_strip_xml_namespaces>(&xml[0]);

  rapidxml::xml_node<>* sst = sharedStrings.first_node("sst");

  unsigned long int n = 0;
  rapidxml::xml_attribute<>* uniqueCount = sst->first_attribute("uniqueCount");
  if (uniqueCount != NULL) { // The count is stored in the file
    n = strtol(uniqueCount->value(), NULL, 10);
  } else { // Count them manually
    for (rapidxml::xml_node<>* string = sst->first_node();
        string; string = string->next_sibling()) {
      n += 1;
    }
  }
  strings_.reserve(n);
  strings_formatted_ = Rcpp::List(n);

  // 18.4.8 si (String Item) [p1725]
  unsigned long int i = 0;
  for (rapidxml::xml_node<>* string = sst->first_node();
      string; string = string->next_sibling()) {
    std::string out;
    parseString(string, out);    // missing strings are treated as empty ""
    strings_.push_back(out);

    Rcpp::List out_df = parseFormattedString(string, styles_);
    strings_formatted_[i] = out_df;
    i += 1;
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
    String namestring(*name);
    String comments_path_string(*comments_path);
    sheets_.emplace_back(
        xlsxsheet(
          namestring,
          xmlstring,
          *this,
          comments_path_string,
          include_blank_cells_)
        );
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
  content_         = CharacterVector(cellcount_, NA_STRING);
  data_type_       = CharacterVector(cellcount_, NA_STRING);
  error_           = CharacterVector(cellcount_, NA_STRING);
  logical_         = LogicalVector(cellcount_,   NA_LOGICAL);
  numeric_         = NumericVector(cellcount_,   NA_REAL);
  date_            = NumericVector(cellcount_,   NA_REAL);
  date_.attr("class") = CharacterVector::create("POSIXct", "POSIXt");
  date_.attr("tzone") = "UTC";
  character_       = CharacterVector(cellcount_, NA_STRING);
  formula_         = CharacterVector(cellcount_, NA_STRING);
  is_array_        = LogicalVector(cellcount_,   false);
  formula_ref_     = CharacterVector(cellcount_, NA_STRING);
  formula_group_   = IntegerVector(cellcount_,   NA_INTEGER);
  comment_         = CharacterVector(cellcount_, NA_STRING);
  character_formatted_ = List(cellcount_);
  height_          = NumericVector(cellcount_,   NA_REAL);
  width_           = NumericVector(cellcount_,   NA_REAL);
  rowOutlineLevel_ = NumericVector(cellcount_,   NA_REAL);
  colOutlineLevel_ = NumericVector(cellcount_,   NA_REAL);
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
    doc.parse<rapidxml::parse_strip_xml_namespaces>(&(*xml)[0]);
    rapidxml::xml_node<>* workbook = doc.first_node("worksheet");
    rapidxml::xml_node<>* sheetData = workbook->first_node("sheetData");
    sheet->parseSheetData(sheetData, i);
    sheet->appendComments(i);
  }

  // Returns a nested data frame of everything, the data frame itself wrapped in
  // a list.

  information_ = List(24);
  information_[0] = sheet_;
  information_[1] = address_;
  information_[2] = row_;
  information_[3] = col_;
  information_[4] = is_blank_;
  information_[5] = content_;
  information_[6] = data_type_;
  information_[7] = error_;
  information_[8] = logical_;
  information_[9] = numeric_;
  information_[10] = date_;
  information_[11] = character_;
  information_[12] = character_formatted_;
  information_[13] = formula_;
  information_[14] = is_array_;
  information_[15] = formula_ref_;
  information_[16] = formula_group_;
  information_[17] = comment_;
  information_[18] = height_;
  information_[19] = width_;
  information_[20] = rowOutlineLevel_;
  information_[21] = colOutlineLevel_;
  information_[22] = style_format_;
  information_[23] = local_format_id_;

  std::vector<std::string> names(24);
  names[0]  = "sheet";
  names[1]  = "address";
  names[2]  = "row";
  names[3]  = "col";
  names[4]  = "is_blank";
  names[5]  = "content";
  names[6]  = "data_type";
  names[7]  = "error";
  names[8]  = "logical";
  names[9]  = "numeric";
  names[10]  = "date";
  names[11] = "character";
  names[12] = "character_formatted";
  names[13] = "formula";
  names[14] = "is_array";
  names[15] = "formula_ref";
  names[16] = "formula_group";
  names[17] = "comment";
  names[18] = "height";
  names[19] = "width";
  names[20] = "row_outline_level";
  names[21] = "col_outline_level";
  names[22] = "style_format";
  names[23] = "local_format_id";

  information_.attr("names") = names;

  // Turn list of vectors into a data frame without checking anything
  int n = Rf_length(information_[0]);
  information_.attr("class") = CharacterVector::create("tbl_df", "tbl", "data.frame");
  information_.attr("row.names") = IntegerVector::create(NA_INTEGER, -n); // Dunno how this works (the -n part)
}
