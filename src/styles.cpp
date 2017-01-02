#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "styles.h"
#include "xf.h"

using namespace Rcpp;

styles::styles(const std::string& path) {
  std::string styles = zip_buffer(path, "xl/styles.xml");
  rapidxml::xml_document<> styles_xml;
  styles_xml.parse<0>(&styles[0]);
  rapidxml::xml_node<>* styleSheet = styles_xml.first_node("styleSheet");

  cacheThemeRgb(path);
  cacheIndexRgb();
  cacheNumFmts(styleSheet);
  cacheCellXfs(styleSheet);
  cacheCellStyleXfs(styleSheet);
}

void styles::cacheThemeRgb(const std::string& path) {
  CharacterVector theme(12, NA_STRING);
  theme_ = theme;
  std::string FF = "FF";
  if (zip_has_file(path, "xl/theme/theme1.xml")) {
    std::string theme1 = zip_buffer(path, "xl/theme/theme1.xml");
    rapidxml::xml_document<> theme1_xml;
    theme1_xml.parse<0>(&theme1[0]);
    rapidxml::xml_node<>* theme = theme1_xml.first_node("a:theme");
    rapidxml::xml_node<>* themeElements = theme->first_node("a:themeElements");
    rapidxml::xml_node<>* clrScheme = themeElements->first_node("a:clrScheme");

    // First, two sysClr nodes in the wrong order
    rapidxml::xml_node<>* color = clrScheme->first_node();
    theme_[1] = FF + color->first_node()->first_attribute("lastClr")->value();
    color = color->next_sibling();
    theme_[0] = FF + color->first_node()->first_attribute("lastClr")->value();

    // Then, two srgbClr nodes in the wrong order
    color = color->next_sibling();
    theme_[3] = FF + color->first_node()->first_attribute("val")->value();
    color = color->next_sibling();
    theme_[2] = FF + color->first_node()->first_attribute("val")->value();

    // Finally, eight more srgbClr nodes in the correct order
    // Can't reuse 'color' here, so use 'nextcolor'
    int i = 4;
    for (rapidxml::xml_node<>* nextcolor = color->next_sibling();
        nextcolor; nextcolor = nextcolor->next_sibling()) {
      theme_[i] = FF + nextcolor->first_node()->first_attribute("val")->value();
      i++;
    }
  }
}

void styles::cacheIndexRgb() {
  CharacterVector index(82, NA_STRING);
  index[0]  = "FF000000";
  index[1]  = "FFFFFFFF";
  index[2]  = "FFFF0000";
  index[3]  = "FF00FF00";
  index[4]  = "FF0000FF";
  index[5]  = "FFFFFF00";
  index[6]  = "FFFF00FF";
  index[7]  = "FF00FFFF";
  index[8]  = "FF000000";
  index[9]  = "FFFFFFFF";
  index[10] = "FFFF0000";
  index[11] = "FF00FF00";
  index[12] = "FF0000FF";
  index[13] = "FFFFFF00";
  index[14] = "FFFF00FF";
  index[15] = "FF00FFFF";
  index[16] = "FF800000";
  index[17] = "FF008000";
  index[18] = "FF000080";
  index[19] = "FF808000";
  index[20] = "FF800080";
  index[21] = "FF008080";
  index[22] = "FFC0C0C0";
  index[23] = "FF808080";
  index[24] = "FF9999FF";
  index[25] = "FF993366";
  index[26] = "FFFFFFCC";
  index[27] = "FFCCFFFF";
  index[28] = "FF660066";
  index[29] = "FFFF8080";
  index[30] = "FF0066CC";
  index[31] = "FFCCCCFF";
  index[32] = "FF000080";
  index[33] = "FFFF00FF";
  index[34] = "FFFFFF00";
  index[35] = "FF00FFFF";
  index[36] = "FF800080";
  index[37] = "FF800000";
  index[38] = "FF008080";
  index[39] = "FF0000FF";
  index[40] = "FF00CCFF";
  index[41] = "FFCCFFFF";
  index[42] = "FFCCFFCC";
  index[43] = "FFFFFF99";
  index[44] = "FF99CCFF";
  index[45] = "FFFF99CC";
  index[46] = "FFCC99FF";
  index[47] = "FFFFCC99";
  index[48] = "FF3366FF";
  index[49] = "FF33CCCC";
  index[50] = "FF99CC00";
  index[51] = "FFFFCC00";
  index[52] = "FFFF9900";
  index[53] = "FFFF6600";
  index[54] = "FF666699";
  index[55] = "FF969696";
  index[56] = "FF003366";
  index[57] = "FF339966";
  index[58] = "FF003300";
  index[59] = "FF333300";
  index[60] = "FF993300";
  index[61] = "FF993366";
  index[62] = "FF333399";
  index[63] = "FF333333";
  index[64] = "FFFFFFFF"; // System foreground
  index[65] = "FF000000"; // System background

  index[81] = "FF000000"; // Undocumented comment text in black

  index_ = index;
}

void styles::cacheNumFmts(rapidxml::xml_node<>* styleSheet) {
  rapidxml::xml_node<>* numFmts = styleSheet->first_node("numFmts");
  // Iterate through custom formats once to get the maximum ID
  int maxId = 49; // stores the max numFmtId encountered so far.
  // If no custom formats are defined, then 49 covers the default formats.
  if (numFmts != NULL) {
    for (rapidxml::xml_node<>* numFmt = numFmts->first_node("numFmt");
        numFmt; numFmt = numFmt->next_sibling()) {
      int id = strtol(numFmt->first_attribute("numFmtId")->value(), NULL, 10);
      maxId = (id > maxId) ? id : maxId;
    }
  }

  // Define vectors the length of the max Id
  CharacterVector formatCodes(maxId + 1, NA_STRING);
  LogicalVector   isDate(maxId + 1, NA_LOGICAL);

  // Populate the default formats
  formatCodes[0]  = "General";
  formatCodes[1]  = "0";
  formatCodes[2]  = "0.00";
  formatCodes[3]  = "#;##0";
  formatCodes[4]  = "#;##0.00";

  formatCodes[9]  = "0%";
  formatCodes[10] = "0.00%";
  formatCodes[11] = "0.00E+00";
  formatCodes[12] = "# ?/?";
  formatCodes[13] = "# ?\?/?\?"; // Backslashes to escape the trigraph
  formatCodes[14] = "mm-dd-yy";
  formatCodes[15] = "d-mmm-yy";
  formatCodes[16] = "d-mmm";
  formatCodes[17] = "mmm-yy";
  formatCodes[18] = "h:mm AM/PM";
  formatCodes[19] = "h:mm:ss AM/PM";
  formatCodes[20] = "h:mm";
  formatCodes[21] = "h:mm:ss";
  formatCodes[22] = "m/d/yy h:mm";

  formatCodes[37] = "#;##0 ;(#;##0)";
  formatCodes[38] = "#;##0 ;[Red](#;##0)";
  formatCodes[39] = "#;##0.00;(#;##0.00)";
  formatCodes[40] = "#;##0.00;[Red](#;##0.00)";

  formatCodes[45] = "mm:ss";
  formatCodes[46] = "[h]:mm:ss";
  formatCodes[47] = "mmss.0";
  formatCodes[48] = "##0.0E+0";
  formatCodes[49] = "@";

  isDate[0]  = false;
  isDate[1]  = false;
  isDate[2]  = false;
  isDate[3]  = false;
  isDate[4]  = false;

  isDate[9]  = false;
  isDate[10] = false;
  isDate[11] = false;
  isDate[12] = false;
  isDate[13] = false;
  isDate[14] = true;
  isDate[15] = true;
  isDate[16] = true;
  isDate[17] = true;
  isDate[18] = true;
  isDate[19] = true;
  isDate[20] = true;
  isDate[21] = true;
  isDate[22] = true;

  isDate[37] = false;
  isDate[38] = false;
  isDate[39] = false;
  isDate[40] = false;

  isDate[45] = true;
  isDate[46] = true;
  isDate[47] = true;
  isDate[48] = false;
  isDate[49] = false;

  // Iterate through custom formats again to get the format codes and to
  // determine which of them are dates.
  if (numFmts != NULL) {
    for (rapidxml::xml_node<>* numFmt = numFmts->first_node("numFmt");
        numFmt; numFmt = numFmt->next_sibling()) {
      int id = strtol(numFmt->first_attribute("numFmtId")->value(), NULL, 10);
      std::string formatCode = numFmt->first_attribute("formatCode")->value();
      formatCodes[id] = formatCode;
      isDate[id] = isDateFormat(formatCode);
    }
  }
  numFmts_ = formatCodes;
  isDate_ = isDate;
}

// Adapted from hadley/readxl
bool styles::isDateFormat(std::string formatCode) {
  for (size_t i = 0; i < formatCode.size(); ++i) {
    switch (formatCode[i]) {
      case 'd':
      case 'm': // 'mm' for minutes
      case 'y':
      case 'h': // 'hh'
      case 's': // 'ss'
        return true;
      default:
        break;
    }
  }
  return false;
}

void styles::cacheCellXfs(rapidxml::xml_node<>* styleSheet) {
  rapidxml::xml_node<>* cellXfs = styleSheet->first_node("cellXfs");
  for (rapidxml::xml_node<>* xf_node = cellXfs->first_node("xf");
      xf_node; xf_node = xf_node->next_sibling()) {
    xf xf(xf_node);
    cellXfs_.push_back(xf); // Inefficient, but there probably won't be many
  }
}

void styles::cacheCellStyleXfs(rapidxml::xml_node<>* styleSheet) {
  rapidxml::xml_node<>* cellStyleXfs = styleSheet->first_node("cellStyleXfs");
  for (rapidxml::xml_node<>* xf_node = cellStyleXfs->first_node("xf");
      xf_node; xf_node = xf_node->next_sibling()) {
    xf xf(xf_node);
    cellStyleXfs_.push_back(xf); // Inefficient, but there probably won't be many
  }

  // Get the names of the styles, if available
  rapidxml::xml_node<>* cellStyles = styleSheet->first_node("cellStyles");
  if (cellStyles != NULL) {
    for (rapidxml::xml_node<>* cellStyle = cellStyles->first_node("cellStyle");
        cellStyle; cellStyle = cellStyle->next_sibling()) {
      // Inefficient, but there probably won't be many
      cellStyles_.push_back(cellStyle->first_attribute("name")->value());
    }
  }
}
