#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxstyles.h"
#include "xf.h"
#include "color.h"
#include "date.h"

using namespace Rcpp;

inline std::string unescape_numFmt(const std::string& s) {
  std::string out;
  out.reserve(s.size());
  bool escaping(false);
  for (size_t i = 0; i < s.size(); i++) {
    if (escaping) {
      escaping = false;
      out.push_back(s[i]);
    } else if (s[i] == '\\') {
      escaping = true;
      // skip
    } else {
      out.push_back(s[i]);
    }
  }
  return out;
}

xlsxstyles::xlsxstyles(const std::string& path) {
  cacheThemeRgb(path);
  cacheIndexedRgb();

  // Try the styles.xml in the file.  If it doesn't define what is needed,
  // then use a default file
  std::string styles1 = zip_buffer(path, "xl/styles.xml");
  rapidxml::xml_document<> styles_xml1;
  styles_xml1.parse<0>(&styles1[0]);
  rapidxml::xml_node<>* styleSheet1 = styles_xml1.first_node("styleSheet");
  rapidxml::xml_node<>* cellStyleXfs = styleSheet1->first_node("cellStyleXfs");
  if (cellStyleXfs != NULL) {
    rapidxml::xml_node<>* styleSheet1 = styles_xml1.first_node("styleSheet");
    cacheNumFmts(styleSheet1);
    cacheCellXfs(styleSheet1);
    cacheCellStyleXfs(styleSheet1);
    cacheFonts(styleSheet1);
    cacheFills(styleSheet1);
    cacheBorders(styleSheet1);
  } else {
    warning("Default styles used (cellStyleXfs is not defined)");
    std::string styles2 = zip_buffer(extdata() + "/default.xlsx", "xl/styles.xml");
    rapidxml::xml_document<> styles_xml2;
    styles_xml2.parse<0>(&styles2[0]);
    rapidxml::xml_node<>* styleSheet2 = styles_xml2.first_node("styleSheet");
    cacheNumFmts(styleSheet2);
    cacheCellXfs(styleSheet2);
    cacheCellStyleXfs(styleSheet2);
    cacheFonts(styleSheet2);
    cacheFills(styleSheet2);
    cacheBorders(styleSheet2);
  }

  applyFormats();
  style_ = zipFormats(style_formats_, true);
  local_ = zipFormats(local_formats_, false);
}

std::string xlsxstyles::rgb_string(rapidxml::xml_node<>* node) {
  rapidxml::xml_node<>* child = node->first_node();
  std::string name = child->name();
  std::string out;
  if (name == "a:sysClr") {
    out = child->first_attribute("lastClr")->value();
  } else { // assume the name is "srgbClr"
    out = child->first_attribute("val")->value();
  }
  return(out);
}

void xlsxstyles::cacheThemeRgb(const std::string& path) {
  theme_name_ =
    CharacterVector::create("background1",
                            "text1",
                            "background2",
                            "text2",
                            "accent1",
                            "accent2",
                            "accent3",
                            "accent4",
                            "accent5",
                            "accent6",
                            "hyperlink",
                            "followed-hyperlink");
  theme_ = CharacterVector(12, NA_STRING);
  std::string FF = "FF";
  if (zip_has_file(path, "xl/theme/theme1.xml")) {
    std::string theme1 = zip_buffer(path, "xl/theme/theme1.xml");
    rapidxml::xml_document<> theme1_xml;
    theme1_xml.parse<0>(&theme1[0]);
    rapidxml::xml_node<>* theme = theme1_xml.first_node("a:theme");
    rapidxml::xml_node<>* themeElements = theme->first_node("a:themeElements");
    rapidxml::xml_node<>* clrScheme = themeElements->first_node("a:clrScheme");

    // First, four nodes in the wrong order
    rapidxml::xml_node<>* color = clrScheme->first_node();
    theme_[1] = FF + rgb_string(color);
    color = color->next_sibling();
    theme_[0] = FF + rgb_string(color);
    color = color->next_sibling();
    theme_[3] = FF + rgb_string(color);
    color = color->next_sibling();
    theme_[2] = FF + rgb_string(color);

    // Then, eight more nodes in the correct order
    // Can't reuse 'color' here, so use 'nextcolor'
    int i = 4;
    for (rapidxml::xml_node<>* nextcolor = color->next_sibling();
        nextcolor; nextcolor = nextcolor->next_sibling()) {
      theme_[i] = FF + rgb_string(nextcolor);
      i++;
    }
  }
}

void xlsxstyles::cacheIndexedRgb() {
  CharacterVector indexed(82, NA_STRING);
  indexed[0]  = "FF000000";
  indexed[1]  = "FFFFFFFF";
  indexed[2]  = "FFFF0000";
  indexed[3]  = "FF00FF00";
  indexed[4]  = "FF0000FF";
  indexed[5]  = "FFFFFF00";
  indexed[6]  = "FFFF00FF";
  indexed[7]  = "FF00FFFF";
  indexed[8]  = "FF000000";
  indexed[9]  = "FFFFFFFF";
  indexed[10] = "FFFF0000";
  indexed[11] = "FF00FF00";
  indexed[12] = "FF0000FF";
  indexed[13] = "FFFFFF00";
  indexed[14] = "FFFF00FF";
  indexed[15] = "FF00FFFF";
  indexed[16] = "FF800000";
  indexed[17] = "FF008000";
  indexed[18] = "FF000080";
  indexed[19] = "FF808000";
  indexed[20] = "FF800080";
  indexed[21] = "FF008080";
  indexed[22] = "FFC0C0C0";
  indexed[23] = "FF808080";
  indexed[24] = "FF9999FF";
  indexed[25] = "FF993366";
  indexed[26] = "FFFFFFCC";
  indexed[27] = "FFCCFFFF";
  indexed[28] = "FF660066";
  indexed[29] = "FFFF8080";
  indexed[30] = "FF0066CC";
  indexed[31] = "FFCCCCFF";
  indexed[32] = "FF000080";
  indexed[33] = "FFFF00FF";
  indexed[34] = "FFFFFF00";
  indexed[35] = "FF00FFFF";
  indexed[36] = "FF800080";
  indexed[37] = "FF800000";
  indexed[38] = "FF008080";
  indexed[39] = "FF0000FF";
  indexed[40] = "FF00CCFF";
  indexed[41] = "FFCCFFFF";
  indexed[42] = "FFCCFFCC";
  indexed[43] = "FFFFFF99";
  indexed[44] = "FF99CCFF";
  indexed[45] = "FFFF99CC";
  indexed[46] = "FFCC99FF";
  indexed[47] = "FFFFCC99";
  indexed[48] = "FF3366FF";
  indexed[49] = "FF33CCCC";
  indexed[50] = "FF99CC00";
  indexed[51] = "FFFFCC00";
  indexed[52] = "FFFF9900";
  indexed[53] = "FFFF6600";
  indexed[54] = "FF666699";
  indexed[55] = "FF969696";
  indexed[56] = "FF003366";
  indexed[57] = "FF339966";
  indexed[58] = "FF003300";
  indexed[59] = "FF333300";
  indexed[60] = "FF993300";
  indexed[61] = "FF993366";
  indexed[62] = "FF333399";
  indexed[63] = "FF333333";
  indexed[64] = "FFFFFFFF"; // System foreground
  indexed[65] = "FF000000"; // System background

  indexed[81] = "FF000000"; // Undocumented comment text in black

  indexed_ = indexed;
}

void xlsxstyles::cacheNumFmts(rapidxml::xml_node<>* styleSheet) {
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
      formatCodes[id] = unescape_numFmt(formatCode);
      isDate[id] = isDateFormat(formatCode);
    }
  }
  numFmts_ = formatCodes;
  isDate_ = isDate;
}

void xlsxstyles::cacheFonts(rapidxml::xml_node<>* styleSheet) {
  rapidxml::xml_node<>* fonts = styleSheet->first_node("fonts");
  for (rapidxml::xml_node<>* font_node = fonts->first_node("font");
      font_node; font_node = font_node->next_sibling()) {
    font font(font_node, this);
    fonts_.push_back(font);
  }
}

void xlsxstyles::cacheFills(rapidxml::xml_node<>* styleSheet) {
  rapidxml::xml_node<>* fills = styleSheet->first_node("fills");
  for (rapidxml::xml_node<>* fill_node = fills->first_node("fill");
      fill_node; fill_node = fill_node->next_sibling()) {
    fill fill(fill_node, this);
    fills_.push_back(fill);
  }
}

void xlsxstyles::cacheBorders(rapidxml::xml_node<>* styleSheet) {
  rapidxml::xml_node<>* borders = styleSheet->first_node("borders");
  for (rapidxml::xml_node<>* border_node = borders->first_node("border");
      border_node; border_node = border_node->next_sibling()) {
    border border(border_node, this);
    borders_.push_back(border);
  }
}

void xlsxstyles::cacheCellXfs(rapidxml::xml_node<>* styleSheet) {
  rapidxml::xml_node<>* cellXfs = styleSheet->first_node("cellXfs");
  for (rapidxml::xml_node<>* xf_node = cellXfs->first_node("xf");
      xf_node; xf_node = xf_node->next_sibling()) {
    xf xf(xf_node);
    cellXfs_.push_back(xf);
  }
}

void xlsxstyles::cacheCellStyleXfs(rapidxml::xml_node<>* styleSheet) {
  rapidxml::xml_node<>* cellStyleXfs = styleSheet->first_node("cellStyleXfs");
  int i(0);
  for (rapidxml::xml_node<>* xf_node = cellStyleXfs->first_node("xf");
      xf_node; xf_node = xf_node->next_sibling()) {
    xf xf(xf_node);
    cellStyleXfs_.push_back(xf);
    ++i;
  }

  // Get the names of the styles, if available
  rapidxml::xml_node<>* cellStyles = styleSheet->first_node("cellStyles");
  if (cellStyles != NULL) {
    // Get the names, which aren't necessarily in xf order
    int index;
    int max_index = 0;
    for (rapidxml::xml_node<>* cellStyle = cellStyles->first_node("cellStyle");
        cellStyle; cellStyle = cellStyle->next_sibling()) {
      index = strtol(cellStyle->first_attribute("xfId")->value(), NULL, 10);
      cellStyles_map_.insert({index, cellStyle->first_attribute("name")->value()});
      if (index > max_index) {
        max_index = index;
      }
    }
    // Sort them
    for (std::map<int, std::string>::iterator i = cellStyles_map_.begin();
        i != cellStyles_map_.end(); i++) {
      cellStyles_.push_back(i->second);
    }
  } else {
    // Gnumeric xlsx files have a single, unnamed style
    cellStyles_.push_back(NA_STRING);
  }
}

void xlsxstyles::clone_color(color& from, colors& to, int& i) {
    to.rgb[i] = from.rgb_;
    to.theme[i] = from.theme_;
    to.indexed[i] = from.indexed_;
    to.tint[i] = from.tint_;
}

List xlsxstyles::list_color(colors& original, bool is_style) {
  if (is_style) {
    original.rgb.attr("names") = cellStyles_;
    original.theme.attr("names") = cellStyles_;
    original.indexed.attr("names") = cellStyles_;
    original.tint.attr("names") = cellStyles_;
  }

  List out = List::create(
      _["rgb"] = original.rgb,
      _["theme"] = original.theme,
      _["indexed"] = original.indexed,
      _["tint"] = original.tint);
  return(out);
}

void xlsxstyles::applyFormats() {
  // Get the normal style
  xf normal = cellStyleXfs_[0];

  style_formats_.emplace_back(xf());
  style_formats_[0].numFmtId_ = normal.numFmtId_;
  style_formats_[0].fontId_ = normal.fontId_;
  style_formats_[0].fillId_ = normal.fillId_;
  style_formats_[0].borderId_ = normal.borderId_;
  style_formats_[0].horizontal_ = normal.horizontal_.get_sexp(); // force copy
  style_formats_[0].vertical_ = normal.vertical_.get_sexp();
  style_formats_[0].wrapText_ = normal.wrapText_;
  style_formats_[0].readingOrder_ = normal.readingOrder_.get_sexp();
  style_formats_[0].indent_ = normal.indent_;
  style_formats_[0].justifyLastLine_ = normal.justifyLastLine_;
  style_formats_[0].shrinkToFit_ = normal.shrinkToFit_;
  style_formats_[0].textRotation_ = normal.textRotation_;
  style_formats_[0].locked_ = normal.locked_;
  style_formats_[0].hidden_ = normal.hidden_;

  // For all other style_formats_, override the normal style if 'apply X' is 1 or NA
  int i(1);
  for(std::vector<xf>::iterator it = cellStyleXfs_.begin() + 1;
      it != cellStyleXfs_.end(); ++it) {
    style_formats_.emplace_back(xf());
    if (it->applyNumberFormat_ != 0) {
      style_formats_[i].numFmtId_ = (it->numFmtId_);
    } else {
      style_formats_[i].numFmtId_ = (normal.numFmtId_);
    }

    if (it->applyFont_ != 0) {
      style_formats_[i].fontId_ = (it->fontId_);
    } else {
      style_formats_[i].fontId_ = (normal.fontId_);
    }

    if (it->applyFill_ != 0) {
      style_formats_[i].fillId_ = (it->fillId_);
    } else {
      style_formats_[i].fillId_ = (normal.fillId_);
    }

    if (it->applyBorder_ != 0) {
      style_formats_[i].borderId_ = (it->borderId_);
    } else {
      style_formats_[i].borderId_ = (normal.borderId_);
    }

    if (it->applyAlignment_ != 0) {
      // Apply the style's style
      style_formats_[i].horizontal_ = it->horizontal_.get_sexp();
      style_formats_[i].vertical_ = it->vertical_.get_sexp();
      style_formats_[i].wrapText_ = it->wrapText_;
      style_formats_[i].readingOrder_ = it->readingOrder_.get_sexp();
      style_formats_[i].indent_ = it->indent_;
      style_formats_[i].justifyLastLine_ = it->justifyLastLine_;
      style_formats_[i].shrinkToFit_ = it->shrinkToFit_;
      style_formats_[i].textRotation_ = it->textRotation_;
    } else {
      // Inherit the 'normal' style
      style_formats_[i].horizontal_ = normal.horizontal_.get_sexp();
      style_formats_[i].vertical_ = normal.vertical_.get_sexp();
      style_formats_[i].wrapText_ = normal.wrapText_;
      style_formats_[i].readingOrder_ = normal.readingOrder_.get_sexp();
      style_formats_[i].indent_ = normal.indent_;
      style_formats_[i].justifyLastLine_ = normal.justifyLastLine_;
      style_formats_[i].shrinkToFit_ = normal.shrinkToFit_;
      style_formats_[i].textRotation_ = normal.textRotation_;
    }

    if (it->applyProtection_ == 1) {
      // Apply the style's style
      // I can't produce this situation in Excel 2016
      style_formats_[i].locked_ = (it->locked_);
      style_formats_[i].hidden_ = (it->hidden_);
    } else {
      // Inherit the 'normal' style
      style_formats_[i].locked_ = (normal.locked_);
      style_formats_[i].hidden_ = (normal.hidden_);
    }
    ++i;
  }

  // Similarly, override local formatting with the local style unless 'apply X' is true
  i = 0;
  for(std::vector<xf>::iterator it = cellXfs_.begin();
      it != cellXfs_.end(); ++it) {
    local_formats_.emplace_back(xf());
    int xfId = it->xfId_; // Use to look up the overall 'style', which is locally modified

    if (it->applyNumberFormat_ == 1 && it->numFmtId_ != NA_INTEGER) {
      local_formats_[i].numFmtId_ = (it->numFmtId_);
    } else { // no known case
      local_formats_[i].numFmtId_ = (style_formats_[xfId].numFmtId_); // # nocov
    }

    if (it->applyFont_ == 1) {
      local_formats_[i].fontId_ = (it->fontId_);
    } else {
      local_formats_[i].fontId_ = (style_formats_[xfId].fontId_);
    }

    if (it->applyFill_ == 1) {
      local_formats_[i].fillId_ = (it->fillId_);
    } else { // no known case
      local_formats_[i].fillId_ = (style_formats_[xfId].fillId_); // # nocov
    }

    if (it->applyBorder_ == 1) {
      local_formats_[i].borderId_ = (it->borderId_);
    } else {
      local_formats_[i].borderId_ = (style_formats_[xfId].borderId_);
    }

    if (it->applyAlignment_ == 1) {
      // Apply the style's style
      local_formats_[i].horizontal_ = it->horizontal_.get_sexp();
      local_formats_[i].vertical_ = it->vertical_.get_sexp();
      local_formats_[i].wrapText_ = it->wrapText_;
      local_formats_[i].readingOrder_ = it->readingOrder_.get_sexp();
      local_formats_[i].indent_ = it->indent_;
      local_formats_[i].justifyLastLine_ = it->justifyLastLine_;
      local_formats_[i].shrinkToFit_ = it->shrinkToFit_;
      local_formats_[i].textRotation_ = it->textRotation_;
    } else {
      // Inherit the 'style_formats_' style
      local_formats_[i].horizontal_ = style_formats_[xfId].horizontal_.get_sexp();
      local_formats_[i].vertical_ = style_formats_[xfId].vertical_.get_sexp();
      local_formats_[i].wrapText_ = style_formats_[xfId].wrapText_;
      local_formats_[i].readingOrder_ = style_formats_[xfId].readingOrder_.get_sexp();
      local_formats_[i].indent_ = style_formats_[xfId].indent_;
      local_formats_[i].justifyLastLine_ = style_formats_[xfId].justifyLastLine_;
      local_formats_[i].shrinkToFit_ = style_formats_[xfId].shrinkToFit_;
      local_formats_[i].textRotation_ = style_formats_[xfId].textRotation_;
    }

    if (it->applyProtection_ == 1) {
      // Apply the style's style
      local_formats_[i].locked_ = (it->locked_);
      local_formats_[i].hidden_ = (it->hidden_);
    } else {
      // Inherit the 'style_formats_' style
      local_formats_[i].locked_ = (style_formats_[xfId].locked_);
      local_formats_[i].hidden_ = (style_formats_[xfId].hidden_);
    }
    ++i;
  }
}

List xlsxstyles::zipFormats(std::vector<xf> styles, bool is_style) {
  int n(styles.size());

  CharacterVector numFmts(n);

  colors fonts_color(n);
  colors fills_patternFill_fgColor(n);
  colors fills_patternFill_bgColor(n);
  colors fills_gradientFill_stop1_color(n);
  colors fills_gradientFill_stop2_color(n);
  colors borders_left_color(n);
  colors borders_right_color(n);
  colors borders_start_color(n); // gnumeric
  colors borders_end_color(n); // gnumeric
  colors borders_top_color(n);
  colors borders_bottom_color(n);
  colors borders_diagonal_color(n);
  colors borders_vertical_color(n);
  colors borders_horizontal_color(n);

  LogicalVector fonts_b(n, NA_LOGICAL);
  LogicalVector fonts_i(n, NA_LOGICAL);
  CharacterVector fonts_u(n, NA_STRING);
  LogicalVector fonts_strike(n, NA_LOGICAL);
  CharacterVector fonts_vertAlign(n, NA_STRING);
  NumericVector fonts_size(n, NA_REAL);
  CharacterVector fonts_name(n, NA_STRING);
  IntegerVector fonts_family(n, NA_INTEGER);
  CharacterVector fonts_scheme(n, NA_STRING);

  CharacterVector fills_patternFill_patternType(n, NA_STRING);
  CharacterVector fills_gradientFill_type(n, NA_STRING);
  IntegerVector fills_gradientFill_degree(n, NA_INTEGER);
  NumericVector fills_gradientFill_left(n, NA_REAL);
  NumericVector fills_gradientFill_right(n, NA_REAL);
  NumericVector fills_gradientFill_top(n, NA_REAL);
  NumericVector fills_gradientFill_bottom(n, NA_REAL);
  NumericVector fills_gradientFill_stop1_position(n, NA_REAL);
  NumericVector fills_gradientFill_stop2_position(n, NA_REAL);

  LogicalVector borders_diagonalDown(n, NA_LOGICAL);
  LogicalVector borders_diagonalUp(n, NA_LOGICAL);
  LogicalVector borders_outline(n, NA_LOGICAL);

  CharacterVector borders_left_style(n, NA_STRING);
  CharacterVector borders_right_style(n, NA_STRING);
  CharacterVector borders_start_style(n, NA_STRING);
  CharacterVector borders_end_style(n, NA_STRING);
  CharacterVector borders_top_style(n, NA_STRING);
  CharacterVector borders_bottom_style(n, NA_STRING);
  CharacterVector borders_diagonal_style(n, NA_STRING);
  CharacterVector borders_vertical_style(n, NA_STRING);
  CharacterVector borders_horizontal_style(n, NA_STRING);

  CharacterVector alignments_horizontal(n, NA_STRING);
  CharacterVector alignments_vertical(n, NA_STRING);
  LogicalVector   alignments_wrapText(n, NA_LOGICAL);
  CharacterVector alignments_readingOrder(n, NA_STRING);
  IntegerVector   alignments_indent(n, NA_INTEGER);
  LogicalVector   alignments_justifyLastLine(n, NA_LOGICAL);
  LogicalVector   alignments_shrinkToFit(n, NA_LOGICAL);
  IntegerVector   alignments_textRotation(n, NA_INTEGER);

  LogicalVector   protections_locked(n, NA_LOGICAL);
  LogicalVector   protections_hidden(n, NA_LOGICAL);

  for(int i = 0; i < styles.size(); ++i) {
    font* font = &fonts_[styles[i].fontId_];
    fill* fill = &fills_[styles[i].fillId_];
    border* border = &borders_[styles[i].borderId_];

    numFmts[i] = numFmts_[styles[i].numFmtId_];

    fonts_b[i] = font->b_;
    fonts_i[i] = font->i_;
    fonts_u[i] = font->u_;
    fonts_strike[i] = font->strike_;
    fonts_vertAlign[i] = font->vertAlign_;
    fonts_size[i] = font->size_;
    clone_color(font->color_, fonts_color, i);
    fonts_name[i] = font->name_;
    fonts_family[i] = font->family_;
    fonts_scheme[i] = font->scheme_;

    clone_color(fill->patternFill_.fgColor_, fills_patternFill_fgColor, i);
    clone_color(fill->patternFill_.bgColor_, fills_patternFill_bgColor, i);
    fills_patternFill_patternType[i] = fill->patternFill_.patternType_;

    fills_gradientFill_type[i] = fill->gradientFill_.type_;
    fills_gradientFill_degree[i] = fill->gradientFill_.degree_;
    fills_gradientFill_left[i] = fill->gradientFill_.left_;
    fills_gradientFill_right[i] = fill->gradientFill_.right_;
    fills_gradientFill_top[i] = fill->gradientFill_.top_;
    fills_gradientFill_bottom[i] = fill->gradientFill_.bottom_;
    fills_gradientFill_stop1_position[i] = fill->gradientFill_.stop1_.position_;
    fills_gradientFill_stop2_position[i] = fill->gradientFill_.stop2_.position_;
    clone_color(fill->gradientFill_.stop1_.color_, fills_gradientFill_stop1_color, i);
    clone_color(fill->gradientFill_.stop2_.color_, fills_gradientFill_stop2_color, i);

    borders_diagonalDown[i] = border->diagonalDown_;
    borders_diagonalUp[i] = border->diagonalUp_;
    borders_outline[i] = border->outline_;

    borders_left_style[i] = border->left_.style_;
    borders_right_style[i] = border->right_.style_;
    borders_start_style[i] = border->start_.style_;
    borders_end_style[i] = border->end_.style_;
    borders_top_style[i] = border->top_.style_;
    borders_bottom_style[i] = border->bottom_.style_;
    borders_diagonal_style[i] = border->diagonal_.style_;
    borders_vertical_style[i] = border->vertical_.style_;
    borders_horizontal_style[i] = border->horizontal_.style_;

    clone_color(border->left_.color_, borders_left_color, i);
    clone_color(border->right_.color_, borders_right_color, i);
    clone_color(border->start_.color_, borders_start_color, i);
    clone_color(border->end_.color_, borders_end_color, i);
    clone_color(border->top_.color_, borders_top_color, i);
    clone_color(border->bottom_.color_, borders_bottom_color, i);
    clone_color(border->diagonal_.color_, borders_diagonal_color, i);
    clone_color(border->vertical_.color_, borders_vertical_color, i);
    clone_color(border->horizontal_.color_, borders_horizontal_color, i);

    alignments_horizontal[i] = styles[i].horizontal_;
    alignments_vertical[i] = styles[i].vertical_;
    alignments_wrapText[i] = styles[i].wrapText_;
    alignments_readingOrder[i] = styles[i].readingOrder_;
    alignments_indent[i] = styles[i].indent_;
    alignments_justifyLastLine[i] = styles[i].justifyLastLine_;
    alignments_shrinkToFit[i] = styles[i].shrinkToFit_;
    alignments_textRotation[i] = styles[i].textRotation_;

    protections_locked[i] = styles[i].locked_;
    protections_hidden[i] = styles[i].hidden_;
  }

  if (is_style) {
    // name all the vectors using the style names
    numFmts.attr("names") = cellStyles_;
    fonts_b.attr("names") = cellStyles_;
    fonts_i.attr("names") = cellStyles_;
    fonts_u.attr("names") = cellStyles_;
    fonts_strike.attr("names") = cellStyles_;
    fonts_vertAlign.attr("names") = cellStyles_;
    fonts_size.attr("names") = cellStyles_;
    fonts_name.attr("names") = cellStyles_;
    fonts_family.attr("names") = cellStyles_;
    fonts_scheme.attr("names") = cellStyles_;
    fills_patternFill_patternType.attr("names") = cellStyles_;
    fills_gradientFill_type.attr("names") = cellStyles_;
    fills_gradientFill_degree.attr("names") = cellStyles_;
    fills_gradientFill_left.attr("names") = cellStyles_;
    fills_gradientFill_right.attr("names") = cellStyles_;
    fills_gradientFill_top.attr("names") = cellStyles_;
    fills_gradientFill_bottom.attr("names") = cellStyles_;
    fills_gradientFill_stop1_position.attr("names") = cellStyles_;
    fills_gradientFill_stop2_position.attr("names") = cellStyles_;
    borders_diagonalDown.attr("names") = cellStyles_;
    borders_diagonalUp.attr("names") = cellStyles_;
    borders_outline.attr("names") = cellStyles_;
    borders_left_style.attr("names") = cellStyles_;
    borders_right_style.attr("names") = cellStyles_;
    borders_start_style.attr("names") = cellStyles_;
    borders_end_style.attr("names") = cellStyles_;
    borders_top_style.attr("names") = cellStyles_;
    borders_bottom_style.attr("names") = cellStyles_;
    borders_diagonal_style.attr("names") = cellStyles_;
    borders_vertical_style.attr("names") = cellStyles_;
    borders_horizontal_style.attr("names") = cellStyles_;
    alignments_horizontal.attr("names") = cellStyles_;
    alignments_vertical.attr("names") = cellStyles_;
    alignments_wrapText.attr("names") = cellStyles_;
    alignments_readingOrder.attr("names") = cellStyles_;
    alignments_indent.attr("names") = cellStyles_;
    alignments_justifyLastLine.attr("names") = cellStyles_;
    alignments_shrinkToFit.attr("names") = cellStyles_;
    alignments_textRotation.attr("names") = cellStyles_;

    protections_locked.attr("names") = cellStyles_;
    protections_hidden.attr("names") = cellStyles_;
  }

  return List::create(
      _["numFmt"] = numFmts,
      _["font"] = List::create(
        _["bold"] = fonts_b,
        _["italic"] = fonts_i,
        _["underline"] = fonts_u,
        _["strike"] = fonts_strike,
        _["vertAlign"] = fonts_vertAlign,
        _["size"] = fonts_size,
        _["color"] = list_color(fonts_color, is_style),
        _["name"] = fonts_name,
        _["family"] = fonts_family,
        _["scheme"] = fonts_scheme),
      _["fill"] = List::create(
        _["patternFill"] = List::create(
          _["fgColor"] = list_color(fills_patternFill_fgColor, is_style),
          _["bgColor"] = list_color(fills_patternFill_bgColor, is_style),
          _["patternType"] = fills_patternFill_patternType),
        _["gradientFill"] = List::create(
          _["type"] = fills_gradientFill_type,
          _["degree"] = fills_gradientFill_degree,
          _["left"] = fills_gradientFill_left,
          _["right"] = fills_gradientFill_right,
          _["top"] = fills_gradientFill_top,
          _["bottom"] = fills_gradientFill_bottom,
          _["stop1"] = List::create(
            _["position"] = fills_gradientFill_stop1_position,
            _["color"] = list_color(fills_gradientFill_stop1_color, is_style)),
          _["stop2"] = List::create(
            _["position"] = fills_gradientFill_stop2_position,
            _["color"] = list_color(fills_gradientFill_stop2_color, is_style)))),
      _["border"] = List::create(
          _["diagonalDown"] = borders_diagonalDown,
          _["diagonalUp"] = borders_diagonalUp,
          _["outline"] = borders_outline,
          _["left"] = List::create(
            _["style"] = borders_left_style,
            _["color"] = list_color(borders_left_color, is_style)),
          _["right"] = List::create(
            _["style"] = borders_right_style,
            _["color"] = list_color(borders_right_color, is_style)),
          _["start"] = List::create(
            _["style"] = borders_start_style,
            _["color"] = list_color(borders_start_color, is_style)),
          _["end"] = List::create(
              _["style"] = borders_end_style,
              _["color"] = list_color(borders_end_color, is_style)),
          _["top"] = List::create(
              _["style"] = borders_top_style,
              _["color"] = list_color(borders_top_color, is_style)),
          _["bottom"] = List::create(
              _["style"] = borders_bottom_style,
              _["color"] = list_color(borders_bottom_color, is_style)),
          _["diagonal"] = List::create(
              _["style"] = borders_diagonal_style,
              _["color"] = list_color(borders_diagonal_color, is_style)),
          _["vertical"] = List::create(
              _["style"] = borders_vertical_style,
              _["color"] = list_color(borders_vertical_color, is_style)),
          _["horizontal"] = List::create(
              _["style"] = borders_horizontal_style,
              _["color"] = list_color(borders_horizontal_color, is_style))),
      _["alignment"] = List::create(
          _["horizontal"] = alignments_horizontal,
          _["vertical"] = alignments_vertical,
          _["wrapText"] = alignments_wrapText,
          _["readingOrder"] = alignments_readingOrder,
          _["indent"] = alignments_indent,
          _["justifyLastLine"] = alignments_justifyLastLine,
          _["shrinkToFit"] = alignments_shrinkToFit,
          _["textRotation"] = alignments_textRotation),
      _["protection"] = List::create(
          _["locked"] = protections_locked,
          _["hidden"] = protections_hidden));
}
