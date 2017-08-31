//[[Rcpp::depends(piton)]] // for pegtl

#ifndef REF_GRAMMAR_
#define REF_GRAMMAR_

#include <pegtl.hpp>
#include "ref.h"

using namespace tao::TAOCPP_PEGTL_NAMESPACE;

namespace xlref
{
  // Three tokens: Ref (A1-style), Text, Other

  // Generic rules
  struct colon : one< ':' > {};
  struct dollar: one< '$' > {};
  struct dot : one< '.' > {};
  struct underscore : one< '_' > {};
  struct backslash : one< '\\' > {};
  struct question : one< '?' > {};
  struct comma: one< ',' > {};
  struct openparen: one< '(' > {};
  struct closeparen: one< ')' > {};

  // Text matches two QuoteD (") and anything between, i.e. character and
  // the surrounding pair of double-quotes.
  struct QuoteD : one< '"' > {};
  struct NotQuoteD : not_one< '"' > {};
  struct DoubleQuotedString : star< sor< seq< QuoteD, QuoteD >,
                                         NotQuoteD > >
  {};
  struct Text : seq< QuoteD, DoubleQuotedString, QuoteD > {};

  // After attempting a Ref, attempt a Text, otherwise consume everything up to
  // the next dollar, comma or parentheses, which are characters that separate
  // other tokens.
  struct sep: sor< dollar, comma, openparen, closeparen > {};
  struct notsep: if_then_else< at< sep >, failure, any > {};
  struct notseps: plus< notsep > {};
  struct Other: sor< sep, notseps > {};
  struct NotRef : sor< Text, Other > {};

  // Anything above 1048576 is not a valid column
  struct MaybeRowToken : rep_min_max< 1, 7, digit > {};
  struct BadRowToken : seq< range< '1', '9' >,
                            range< '0', '9' >,
                            range< '4', '9' >,
                            range< '8', '9' >,
                            range< '5', '9' >,
                            range< '7', '9' >,
                            range< '7', '9' > >
  {};
  struct RowToken : seq< not_at< BadRowToken >, MaybeRowToken > {};

  // Anything above XFD is not a valid column
  struct MaybeColToken : rep_min_max< 1, 3, upper > {};
  struct BadColToken : seq< range< 'X', 'Z' >,
                            range< 'F', 'Z' >,
                            range< 'E', 'Z' > >
  {};
  struct ColToken : seq< not_at< BadColToken >, MaybeColToken > {};

  struct OptDollar : opt< dollar > {};
  struct OptRowToken : seq< OptDollar, RowToken > {};
  struct OptColToken : seq< OptDollar, ColToken > {};

  // Name as in named formula, as well as worksheet names
  // Start with a letter or underscore, continue with word character (letters,
  // numbers and underscore), dot or question mark
  // * first character: [\p{L}\\_]
  // * subsequent characters: [\w\\_\.\?]
  struct NameStartCharacter : sor< alpha, underscore, backslash > {};
  struct NameValidCharacter
    : sor< NameStartCharacter, digit, dot, question >
  {};

  // Attempt to match addresses in this order
  // A:A
  // A1
  // A1:A2
  // 1:1
  struct Ref :
    seq< OptDollar,
         sor< seq< ColToken,
                   if_then_else< colon,
                                 OptColToken,                    // A:A
                                 seq< OptRowToken,               // A1
                                      opt< colon,
                                           OptColToken,
                                           OptRowToken  > > > >, // A1:A2
              seq< RowToken,
                   colon,
                   OptRowToken > >,
         not_at< sor< NameValidCharacter,        // not e.g. A1A or E09904.2!A1
                      disable< openparen > > > > // not e.g. LOG10()
  {};

  // Overall parsing rule
  struct root : seq< opt< Ref >,
                     star< seq< NotRef,
                           opt< Ref > > > > {};

  // Class template for user-defined actions that does
  // nothing by default.
  template<typename Rule>
    struct tokenize : nothing<Rule> {};

  // Specialisations of the user-defined action to do something when a rule
  // succeeds; is called with the portion of the input that matched the rule.

  template<> struct tokenize< Ref >
  {
    template< typename Input >
      static void apply( const Input & in,
                         std::vector<token_type> & types,
                         std::vector<std::string> & tokens,
                         std::vector<ref> & references)
      {
        types.push_back(token_type::REF);
        ref reference(in.string());
        references.push_back(reference);
      }
  };

  template<> struct tokenize< Text >
  {
    template< typename Input >
      static void apply( const Input & in,
                         std::vector<token_type> & types,
                         std::vector<std::string> & tokens,
                         std::vector<ref> & references)
      {
        types.push_back(token_type::TEXT);
        tokens.push_back(in.string());
      }
  };

  template<> struct tokenize< Other >
  {
    template< typename Input >
      static void apply( const Input & in,
                         std::vector<token_type> & types,
                         std::vector<std::string> & tokens,
                         std::vector<ref> & references)
      {
        types.push_back(token_type::OTHER);
        tokens.push_back(in.string());
      }
  };

} // xlref

#endif
