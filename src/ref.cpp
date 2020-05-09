#include "ref.h"
#include <string>

ref::ref(const std::string& text): text_(text) {

  fixcol1_ = false;
  col1_ = 0;
  fixrow1_ = false;
  row1_ = 0;
  colon_ = false;
  fixcol2_ = false;
  col2_ = 0;
  fixrow2_ = false;
  row2_ = 0;

  std::string::const_iterator iter = text_.begin();

  // Check for a $ fix symbol
  fixcol1_ = *iter == '$';
  if (fixcol1_) ++iter;

  if (*iter >= 'A' && *iter <= 'Z') { // If it's a column
    col1_ = 0;
    for(; (*iter >= 'A' && *iter <= 'Z'); ++iter) {
      col1_ = 26 * col1_ + (*iter - 'A' + 1); // Multiply existing col by 26 and add new col
    }
  }

  // Check for a $ fix symbol
  fixrow1_ = *iter == '$';
  if (fixrow1_) ++iter;

  if (*iter >= '0' && *iter <= '9') { // If it's a row
    row1_ = 0;
    for(; (*iter >= '0' && *iter <= '9'); ++iter) {
      row1_ = 10 * row1_ + (*iter - '0'); // Multiply existing row by 10 and add new row
    }
  }

  // If there's a : range symbol then parse the other side of the range
  colon_ = *iter == ':';
  if (colon_) {
    ++iter;
    // Check for a $ fix symbol
    fixcol2_ = *iter == '$';
    if (fixcol2_) ++iter;

    if (*iter >= 'A' && *iter <= 'Z') { // If it's a column
      col2_ = 0;
      for(; (*iter >= 'A' && *iter <= 'Z'); ++iter) {
        col2_ = 26 * col2_ + (*iter - 'A' + 1); // Multiply existing col by 26 and add new col
      }
    }

    // Check for a $ fix symbol
    fixrow2_ = *iter == '$';
    if (fixrow2_) ++iter;

    if (*iter >= '0' && *iter <= '9') { // If it's a row
      row2_ = 0;
      for(; (*iter >= '0' && *iter <= '9'); ++iter) {
        row2_ = 10 * row2_ + (*iter - '0'); // Multiply existing row by 10 and add new row
      }
    }
  }
}

std::string ref::offset(int& rows, int& cols) const {
  std::string out;

  if (fixcol1_) {
    out = out + '$';
    if (col1_) out = out + int_to_alpha(col1_);
  } else {
    if (col1_) out = out + int_to_alpha(col1_ + cols);
  }

  if (fixrow1_) {
    out = out + '$';
    if (row1_) out = out + std::to_string(row1_);
  } else {
    if (row1_) out = out + std::to_string(row1_ + rows);
  }

  if (colon_) out = out + ':';

  if (fixcol2_) {
    out = out + '$';
    if (col2_) out = out + int_to_alpha(col2_);
  } else {
    if (col2_) out = out + int_to_alpha(col2_ + cols);
  }

  if (fixrow2_) {
    out = out + '$';
    if (row2_) out = out + std::to_string(row2_);
  } else {
    if (row2_) out = out + std::to_string(row2_ + rows);
  }

  return out;
}

// Adapted from http://blog.bigsmoke.us/2010/12/08/cpp-base10-to-base26
std::string ref::int_to_alpha(int i) const
{
    std::string out;
    int dividend = i;
    int modulo;

    while (dividend > 0)
    {
        modulo = (dividend - 1) % 26;
        out = (char)(65 + modulo) + out;
        dividend = (int)((dividend - modulo) / 26);
    }

    return out;
}
