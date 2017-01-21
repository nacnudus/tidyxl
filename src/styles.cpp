#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "styles.h"
#include "xf.h"
#include "color.h"

using namespace Rcpp;

styles::styles(const std::string& path) {
  std::string styles = zip_buffer(path, "xl/styles.xml");
  rapidxml::xml_document<> styles_xml;
  styles_xml.parse<0>(&styles[0]);
  rapidxml::xml_node<>* styleSheet = styles_xml.first_node("styleSheet");

  cacheThemeRgb(path);
  cacheIndexedRgb();
  cacheNumFmts(styleSheet);
  cacheCellXfs(styleSheet);
  cacheCellStyleXfs(styleSheet);
  cacheFonts(styleSheet);
  cacheFills(styleSheet);
  cacheBorders(styleSheet);

  applyFormats();
  style_ = zipFormats(style_formats_);
  local_ = zipFormats(local_formats_);
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

void styles::cacheIndexedRgb() {
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

void styles::cacheFonts(rapidxml::xml_node<>* styleSheet) {
  rapidxml::xml_node<>* fonts = styleSheet->first_node("fonts");
  for (rapidxml::xml_node<>* font_node = fonts->first_node("font");
      font_node; font_node = font_node->next_sibling()) {
    font font(font_node, this);
    fonts_.push_back(font); // Inefficient, but there probably won't be many
  }
}

void styles::cacheFills(rapidxml::xml_node<>* styleSheet) {
  rapidxml::xml_node<>* fills = styleSheet->first_node("fills");
  for (rapidxml::xml_node<>* fill_node = fills->first_node("fill");
      fill_node; fill_node = fill_node->next_sibling()) {
    fill fill(fill_node, this);
    fills_.push_back(fill); // Inefficient, but there probably won't be many
  }
}

void styles::cacheBorders(rapidxml::xml_node<>* styleSheet) {
  rapidxml::xml_node<>* borders = styleSheet->first_node("borders");
  for (rapidxml::xml_node<>* border_node = borders->first_node("border");
      border_node; border_node = border_node->next_sibling()) {
    border border(border_node, this);
    borders_.push_back(border); // Inefficient, but there probably won't be many
  }
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
  CharacterVector stylenames;
  if (cellStyles != NULL) {
    // Get the names, which aren't necessarily in xf order
    for (rapidxml::xml_node<>* cellStyle = cellStyles->first_node("cellStyle");
        cellStyle; cellStyle = cellStyle->next_sibling()) {
      // Inefficient, but there probably won't be many
      stylenames.push_back(cellStyle->first_attribute("name")->value());
    }
    // Go through again, getting the xfId and using that to push the names into
    // the global vector, in order this time.
    int index;
    for (rapidxml::xml_node<>* cellStyle = cellStyles->first_node("cellStyle");
        cellStyle; cellStyle = cellStyle->next_sibling()) {
      index = strtol(cellStyle->first_attribute("xfId")->value(), NULL, 10);
      cellStyles_.push_back(stylenames[index]);
    }
  }
}

void styles::clone_color(color& from, colors& to) {
    to.rgb.push_back(from.rgb_);
    to.theme.push_back(from.theme_[0]);
    to.indexed.push_back(from.indexed_[0]);
    to.tint.push_back(from.tint_[0]);
}

List styles::list_color(colors& original) {
  List out = List::create(
      _["rgb"] = original.rgb,
      _["theme"] = original.theme,
      _["indexed"] = original.indexed,
      _["tint"] = original.tint);
  return(out);
}

void styles::applyFormats() {
  // Get the normal style
  xf normal = cellStyleXfs_[0];
  style_formats_.numFmtId_.push_back(normal.numFmtId_[0]);
  style_formats_.fontId_.push_back(normal.fontId_[0]);
  style_formats_.fillId_.push_back(normal.fillId_[0]);
  style_formats_.borderId_.push_back(normal.borderId_[0]);
  style_formats_.horizontal_.push_back(normal.horizontal_[0]);
  style_formats_.vertical_.push_back(normal.vertical_[0]);
  style_formats_.wrapText_.push_back(normal.wrapText_[0]);
  style_formats_.readingOrder_.push_back(normal.readingOrder_[0]);
  style_formats_.indent_.push_back(normal.indent_[0]);
  style_formats_.justifyLastLine_.push_back(normal.justifyLastLine_[0]);
  style_formats_.shrinkToFit_.push_back(normal.shrinkToFit_[0]);
  style_formats_.textRotation_.push_back(normal.textRotation_[0]);
  style_formats_.locked_.push_back(normal.locked_[0]);
  style_formats_.hidden_.push_back(normal.hidden_[0]);

  // For all other style_formats_, override the normal style if 'apply X' is 1 or NA
  for(std::vector<xf>::iterator it = cellStyleXfs_.begin() + 1;
      it != cellStyleXfs_.end(); ++it) {
    if (it->applyNumberFormat_[0] != 0) {
      style_formats_.numFmtId_.push_back(it->numFmtId_[0]);
    } else {
      style_formats_.numFmtId_.push_back(normal.numFmtId_[0]);
    }

    if (it->applyFont_[0] != 0) {
      style_formats_.fontId_.push_back(it->fontId_[0]);
    } else {
      style_formats_.fontId_.push_back(normal.fontId_[0]);
    }

    if (it->applyFill_[0] != 0) {
      style_formats_.fillId_.push_back(it->fillId_[0]);
    } else {
      style_formats_.fillId_.push_back(normal.fillId_[0]);
    }

    if (it->applyBorder_[0] != 0) {
      style_formats_.borderId_.push_back(it->borderId_[0]);
    } else {
      style_formats_.borderId_.push_back(normal.borderId_[0]);
    }

    if (it->applyAlignment_[0] != 0) {
      // Apply the style's style
      style_formats_.horizontal_.push_back(it->horizontal_[0]);
      style_formats_.vertical_.push_back(it->vertical_[0]);
      style_formats_.wrapText_.push_back(it->wrapText_[0]);
      style_formats_.readingOrder_.push_back(it->readingOrder_[0]);
      style_formats_.indent_.push_back(it->indent_[0]);
      style_formats_.justifyLastLine_.push_back(it->justifyLastLine_[0]);
      style_formats_.shrinkToFit_.push_back(it->shrinkToFit_[0]);
      style_formats_.textRotation_.push_back(it->textRotation_[0]);
    } else {
      // Inherit the 'normal' style
      style_formats_.horizontal_.push_back(normal.horizontal_[0]);
      style_formats_.vertical_.push_back(normal.vertical_[0]);
      style_formats_.wrapText_.push_back(normal.wrapText_[0]);
      style_formats_.readingOrder_.push_back(normal.readingOrder_[0]);
      style_formats_.indent_.push_back(normal.indent_[0]);
      style_formats_.justifyLastLine_.push_back(normal.justifyLastLine_[0]);
      style_formats_.shrinkToFit_.push_back(normal.shrinkToFit_[0]);
      style_formats_.textRotation_.push_back(normal.textRotation_[0]);
    }

    if (it->applyProtection_[0] == 1) {
      // Apply the style's style
      style_formats_.locked_.push_back(it->locked_[0]);
      style_formats_.hidden_.push_back(it->hidden_[0]);
    } else {
      // Inherit the 'normal' style
      style_formats_.locked_.push_back(normal.locked_[0]);
      style_formats_.hidden_.push_back(normal.hidden_[0]);
    }
  }

  // Similarly, override local formatting with the local style unless 'apply X' is true
  for(std::vector<xf>::iterator it = cellXfs_.begin();
      it != cellXfs_.end(); ++it) {
    int xfId = it->xfId_[0]; // Use to look up the overall 'style', which is locally modified

    if (it->applyNumberFormat_[0] == 1) {
      local_formats_.numFmtId_.push_back(it->numFmtId_[0]);
    } else {
      local_formats_.numFmtId_.push_back(style_formats_.numFmtId_[xfId]);
    }

    if (it->applyFont_[0] == 1) {
      local_formats_.fontId_.push_back(it->fontId_[0]);
    } else {
      local_formats_.fontId_.push_back(style_formats_.fontId_[xfId]);
    }

    if (it->applyFill_[0] == 1) {
      local_formats_.fillId_.push_back(it->fillId_[0]);
    } else {
      local_formats_.fillId_.push_back(style_formats_.fillId_[xfId]);
    }

    if (it->applyBorder_[0] == 1) {
      local_formats_.borderId_.push_back(it->borderId_[0]);
    } else {
      local_formats_.borderId_.push_back(style_formats_.borderId_[xfId]);
    }

    if (it->applyAlignment_[0] == 1) {
      // Apply the style's style
      local_formats_.horizontal_.push_back(it->horizontal_[0]);
      local_formats_.vertical_.push_back(it->vertical_[0]);
      local_formats_.wrapText_.push_back(it->wrapText_[0]);
      local_formats_.readingOrder_.push_back(it->readingOrder_[0]);
      local_formats_.indent_.push_back(it->indent_[0]);
      local_formats_.justifyLastLine_.push_back(it->justifyLastLine_[0]);
      local_formats_.shrinkToFit_.push_back(it->shrinkToFit_[0]);
      local_formats_.textRotation_.push_back(it->textRotation_[0]);
    } else {
      // Inherit the 'style_formats_' style
      local_formats_.horizontal_.push_back(style_formats_.horizontal_[xfId]);
       local_formats_.vertical_.push_back(style_formats_.vertical_[xfId]);
      local_formats_.wrapText_.push_back(style_formats_.wrapText_[xfId]);
      local_formats_.readingOrder_.push_back(style_formats_.readingOrder_[xfId]);
      local_formats_.indent_.push_back(style_formats_.indent_[xfId]);
      local_formats_.justifyLastLine_.push_back(style_formats_.justifyLastLine_[xfId]);
      local_formats_.shrinkToFit_.push_back(style_formats_.shrinkToFit_[xfId]);
      local_formats_.textRotation_.push_back(style_formats_.textRotation_[xfId]);
    }

    if (it->applyProtection_[0] == 1) {
      // Apply the style's style
      local_formats_.locked_.push_back(it->locked_[0]);
      local_formats_.hidden_.push_back(it->hidden_[0]);
    } else {
      // Inherit the 'style_formats_' style
      local_formats_.locked_.push_back(style_formats_.locked_[xfId]);
      local_formats_.hidden_.push_back(style_formats_.hidden_[xfId]);
    }
  }
}

List styles::zipFormats(xf styles) {
  CharacterVector style_numFmt;
  font style_font;
  fill style_fill;
  border style_border;

  CharacterVector rgb;
  IntegerVector theme;
  IntegerVector indexed;
  NumericVector tint;

  colors style_font_color;
  colors style_fill_patternFill_fgColor;
  colors style_fill_patternFill_bgColor;
  colors style_fill_gradientFill_color1;
  colors style_fill_gradientFill_color2;
  colors style_border_left_color;
  colors style_border_right_color;
  colors style_border_start_color; // gnumeric
  colors style_border_end_color; // gnumeric
  colors style_border_top_color;
  colors style_border_bottom_color;
  colors style_border_diagonal_color;
  colors style_border_vertical_color;
  colors style_border_horizontal_color;

  CharacterVector style_font_u;
  CharacterVector style_font_vertAlign;
  CharacterVector style_font_name;
  CharacterVector style_font_scheme;

  CharacterVector style_fill_patternFill_patternType;
  CharacterVector style_fill_gradientFill_type;

  CharacterVector style_border_left_style;
  CharacterVector style_border_right_style;
  CharacterVector style_border_start_style;
  CharacterVector style_border_end_style;
  CharacterVector style_border_top_style;
  CharacterVector style_border_bottom_style;
  CharacterVector style_border_diagonal_style;
  CharacterVector style_border_vertical_style;
  CharacterVector style_border_horizontal_style;

  for(int i = 0; i < styles.numFmtId_.size(); ++i) {
    font* font = &fonts_[styles.fontId_[i]];
    fill* fill = &fills_[styles.fillId_[i]];
    border* border = &borders_[styles.borderId_[i]];

    style_numFmt.push_back(numFmts_[styles.numFmtId_[i]]);

    style_font.b_.push_back(font->b_[0]);
    style_font.i_.push_back(font->i_[0]);
    style_font_u.push_back(font->u_);
    style_font.strike_.push_back(font->strike_[0]);
    style_font_vertAlign.push_back(font->vertAlign_);
    style_font.size_.push_back(font->size_[0]);
    clone_color(font->color_, style_font_color);
    style_font_name.push_back(font->name_);
    style_font.family_.push_back(font->family_[0]);
    style_font_scheme.push_back(font->scheme_);

    clone_color(fill->patternFill_.fgColor_, style_fill_patternFill_fgColor);
    clone_color(fill->patternFill_.bgColor_, style_fill_patternFill_bgColor);
    style_fill_patternFill_patternType.push_back(fill->patternFill_.patternType_);

    style_fill_gradientFill_type.push_back(fill->gradientFill_.type_);
    style_fill.gradientFill_.degree_.push_back(fill->gradientFill_.degree_[0]);
    style_fill.gradientFill_.left_.push_back(fill->gradientFill_.left_[0]);
    style_fill.gradientFill_.right_.push_back(fill->gradientFill_.right_[0]);
    style_fill.gradientFill_.top_.push_back(fill->gradientFill_.top_[0]);
    style_fill.gradientFill_.bottom_.push_back(fill->gradientFill_.bottom_[0]);
    clone_color(fill->gradientFill_.color1_, style_fill_gradientFill_color1);
    clone_color(fill->gradientFill_.color2_, style_fill_gradientFill_color2);

    style_border.diagonalDown_.push_back(border->diagonalDown_[0]);
    style_border.diagonalUp_.push_back(border->diagonalUp_[0]);
    style_border.outline_.push_back(border->outline_[0]);

    style_border_left_style.push_back(border->left_.style_);
    style_border_right_style.push_back(border->right_.style_);
    style_border_start_style.push_back(border->start_.style_);
    style_border_end_style.push_back(border->end_.style_);
    style_border_top_style.push_back(border->top_.style_);
    style_border_bottom_style.push_back(border->bottom_.style_);
    style_border_diagonal_style.push_back(border->diagonal_.style_);
    style_border_vertical_style.push_back(border->vertical_.style_);
    style_border_horizontal_style.push_back(border->horizontal_.style_);

    clone_color(border->left_.color_, style_border_left_color);
    clone_color(border->right_.color_, style_border_right_color);
    clone_color(border->start_.color_, style_border_start_color);
    clone_color(border->end_.color_, style_border_end_color);
    clone_color(border->top_.color_, style_border_top_color);
    clone_color(border->bottom_.color_, style_border_bottom_color);
    clone_color(border->diagonal_.color_, style_border_diagonal_color);
    clone_color(border->vertical_.color_, style_border_vertical_color);
    clone_color(border->horizontal_.color_, style_border_horizontal_color);
  }

  return List::create(
      _["numFmt"] = style_numFmt,
      _["font"] = List::create(
        _["bold"] = style_font.b_,
        _["italic"] = style_font.i_,
        _["underline"] = style_font_u,
        _["strike"] = style_font.strike_,
        _["vertAlign"] = style_font_vertAlign,
        _["size"] = style_font.size_,
        _["color"] = list_color(style_font_color),
        _["name"] = style_font_name,
        _["family"] = style_font.family_,
        _["scheme"] = style_font_scheme),
      _["fill"] = List::create(
        _["patternFill"] = List::create(
          _["fgColor"] = list_color(style_fill_patternFill_fgColor),
          _["bgColor"] = list_color(style_fill_patternFill_bgColor),
          _["patternType"] = style_fill_patternFill_patternType),
        _["gradientFill"] = List::create(
          _["type"] = style_fill_gradientFill_type,
          _["degree_"] = style_fill.gradientFill_.degree_,
          _["left"] = style_fill.gradientFill_.left_,
          _["right"] = style_fill.gradientFill_.right_,
          _["top"] = style_fill.gradientFill_.top_,
          _["bottom"] = style_fill.gradientFill_.bottom_,
          _["color1"] = list_color(style_fill_gradientFill_color1),
          _["color2"] = list_color(style_fill_gradientFill_color2))),
      _["border"] = List::create(
          _["diagonalDown"] = style_border.diagonalDown_,
          _["diagonalUp"] = style_border.diagonalUp_,
          _["outline"] = style_border.outline_,
          _["left"] = List::create(
            _["style"] = style_border_left_style,
            _["color"] = list_color(style_border_left_color)),
          _["right"] = List::create(
            _["style"] = style_border_right_style,
            _["color"] = list_color(style_border_right_color)),
          _["start"] = List::create(
            _["style"] = style_border_start_style,
            _["color"] = list_color(style_border_start_color)),
          _["end"] = List::create(
              _["style"] = style_border_end_style,
              _["color"] = list_color(style_border_end_color)),
          _["top"] = List::create(
              _["style"] = style_border_top_style,
              _["color"] = list_color(style_border_top_color)),
          _["bottom"] = List::create(
              _["style"] = style_border_bottom_style,
              _["color"] = list_color(style_border_bottom_color)),
          _["diagonal"] = List::create(
              _["style"] = style_border_diagonal_style,
              _["color"] = list_color(style_border_diagonal_color)),
          _["vertical"] = List::create(
              _["style"] = style_border_vertical_style,
              _["color"] = list_color(style_border_vertical_color)),
          _["horizontal"] = List::create(
              _["style"] = style_border_horizontal_style,
              _["color"] = list_color(style_border_horizontal_color)))
                  );
}
