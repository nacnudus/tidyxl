//[[Rcpp::depends(piton)]] // for pegtl

#ifndef REF_GRAMMAR_
#define REF_GRAMMAR_

#include <pegtl.hpp>
#include "ref.h"

using namespace tao::TAOCPP_PEGTL_NAMESPACE;

namespace xlref
{
  // Three tokens: Ref (A1-style), Text, Other

  // Text matches two QuoteD (") and anything between, i.e. character and
  // the surrounding pair of double-quotes.
  struct QuoteD : one< '"' > {};
  struct NotQuoteD : not_one< '"' > {};
  struct DoubleQuotedString : star< sor< seq< QuoteD, QuoteD >,
                                         NotQuoteD > >
  {};
  struct Text : if_must< QuoteD, DoubleQuotedString, QuoteD > {};

  // After attempting a Ref, consume sequences of ref-valid characters, or
  // sequences of non-ref-valid characters ( not quoted strings), then try
  // again.
  struct Alnum : sor< ranges< 'a', 'z', 'A', 'Z', '0', '9' >,
                      one< '$' > > {};
  struct AlnumQuoteD : sor< ranges< 'a', 'z', 'A', 'Z', '0', '9' >,
                            one< '$' >,
                            one< '"' > > {};
  struct NotAlnumQuoteD : if_then_else< at< AlnumQuoteD >, failure, any > {};
  struct Other: sor< plus< Alnum >, plus< NotAlnumQuoteD > > {};
  struct NotRef : sor< Text, Other > {};

  struct dollar : one< '$' > {};
  struct OptDollar : opt< dollar > {};

  struct RowToken : rep_min_max< 1, 7, digit > {};

  // Anything above XFD is not a valid column
  struct MaybeColToken : rep_min_max< 1, 3, upper > {};
  struct BadColToken : seq< range< 'X', 'Z' >,
                            range< 'F', 'Z' >,
                            range< 'E', 'Z' > > {};
  struct ColToken : seq< not_at< BadColToken >, MaybeColToken > {};

  struct OptRowToken : seq< OptDollar, RowToken > {};
  struct OptColToken : seq< OptDollar, ColToken > {};

  // Attempt to match addresses in this order
  // A:A
  // A1
  // A1:A2
  // 1:1
  struct colon : one< ':' > {};
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
         not_at< Alnum >,                                        // not e.g. A1A
         not_at< one< '(' > > > {};                              // not e.g. LOG10()

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
