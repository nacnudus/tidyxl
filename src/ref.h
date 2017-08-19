#ifndef REF_
#define REF_

#include <string>

enum class token_type {REF, TEXT, OTHER};

class ref
{
  public:

    std::string text_;

    // All the potential parts of a reference $A$1:$B$2
    bool fixcol1_;
    int col1_;
    bool fixrow1_;
    int row1_;
    bool colon_;
    bool fixcol2_;
    int col2_;
    bool fixrow2_;
    int row2_;

    ref(const std::string& text); // parse text into fix/row/col/colon variables

    virtual std::string offset(int& rows, int& cols) const; // return offsetted address

  private:
    std::string int_to_alpha(int i) const;
};

#endif
