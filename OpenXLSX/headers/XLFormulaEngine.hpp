#ifndef OPENXLSX_XLFORMULAENGINE_HPP
#define OPENXLSX_XLFORMULAENGINE_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

// ===== Standard Library ===== //
#include <functional>
#include <memory>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

// ===== GSL ===== //
#include <gsl/gsl>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCellValue.hpp"

// Forward declare XLWorksheet so callers can use makeResolver without pulling the full header.
namespace OpenXLSX { class XLWorksheet; }

namespace OpenXLSX
{
    // =========================================================================
    // Lexer
    // =========================================================================

    /**
     * @brief Kinds of tokens produced by the lexer.
     */
    enum class XLTokenKind : uint8_t
    {
        Number,     ///< Numeric literal (integer or float)
        String,     ///< Quoted string literal  "hello"
        Bool,       ///< TRUE or FALSE keyword
        CellRef,    ///< Cell reference  A1, $B$2, Sheet1!C3
        Ident,      ///< Function name or named range
        Plus,       ///< +
        Minus,      ///< -
        Star,       ///< *
        Slash,      ///< /
        Caret,      ///< ^  (power)
        Percent,    ///< %
        Amp,        ///< &  (string concat)
        Eq,         ///< =
        NEq,        ///< <>
        Lt,         ///< <
        Le,         ///< <=
        Gt,         ///< >
        Ge,         ///< >=
        LParen,     ///< (
        RParen,     ///< )
        Comma,      ///< ,
        Semicolon,  ///< ;  (alternative argument separator in some locales)
        Colon,      ///< :  (used inside range references parsed by lexer)
        End,        ///< Sentinel – end of input
        Error       ///< Unrecognised character
    };

    /**
     * @brief A single lexical token from a formula string.
     */
    struct OPENXLSX_EXPORT XLToken
    {
        XLTokenKind kind{XLTokenKind::Error};
        std::string text;   ///< Raw text of the token (number string, identifier, …)
        double      number{0.0}; ///< Pre-parsed numeric value when kind == Number
        bool        boolean{false}; ///< Pre-parsed bool when kind == Bool
    };

    /**
     * @brief Tokenises a raw Excel formula string (with or without leading '=').
     * @details Hand-written character scanner; zero heap allocations during scanning.
     *          Range references like A1:B10 are emitted as a *single* CellRef token
     *          whose text is "A1:B10".
     */
    class OPENXLSX_EXPORT XLFormulaLexer
    {
    public:
        /**
         * @brief Tokenise the given formula.
         * @param formula The formula string (may start with '=').
         * @return Ordered vector of tokens, last element always XLTokenKind::End.
         */
        static std::vector<XLToken> tokenize(std::string_view formula);

    private:
        // Stateless helper – everything runs in tokenize().
    };

    // =========================================================================
    // AST
    // =========================================================================

    /** @brief Kind discriminator for AST nodes. */
    enum class XLNodeKind : uint8_t
    {
        Number,
        StringLit,
        BoolLit,
        CellRef,    ///< e.g. "A1" or "Sheet1!A1"
        Range,      ///< e.g. "A1:B10" – evaluated lazily by the resolver
        BinOp,
        UnaryOp,
        FuncCall,
        ErrorLit    ///< #NAME?, etc. – propagated as-is
    };

    /**
     * @brief Polymorphic AST node.  Uses a tagged-union approach with `std::unique_ptr` children.
     * @details All fields are public for simplicity; the AST is an internal implementation detail.
     */
    struct OPENXLSX_EXPORT XLASTNode
    {
        XLNodeKind kind;

        // ---- Leaf payloads ----
        double      number{0.0};
        std::string text;       ///< string literal, cell-ref text, range text, identifier, error text
        bool        boolean{false};
        XLTokenKind op{XLTokenKind::Error}; ///< operator for BinOp / UnaryOp

        // ---- Children ----
        std::vector<std::unique_ptr<XLASTNode>> children; ///< operands or function arguments

        explicit XLASTNode(XLNodeKind k) : kind(k) {}

        // Non-copyable due to unique_ptr children; movable.
        XLASTNode(const XLASTNode&)            = delete;
        XLASTNode& operator=(const XLASTNode&) = delete;
        XLASTNode(XLASTNode&&)                 = default;
        XLASTNode& operator=(XLASTNode&&)      = default;
    };

    // =========================================================================
    // Parser
    // =========================================================================

    /**
     * @brief Recursive-descent (Pratt) parser that converts a token stream into an AST.
     */
    class OPENXLSX_EXPORT XLFormulaParser
    {
    public:
        /**
         * @brief Parse the token list produced by XLFormulaLexer::tokenize().
         * @param tokens Span over the token vector (including sentinel End token).
         * @return Root AST node.
         * @throws XLFormulaError on syntax error.
         */
        static std::unique_ptr<XLASTNode> parse(gsl::span<const XLToken> tokens);

    private:
        // All state is local to the recursive parse calls below.
        struct ParseContext
        {
            gsl::span<const XLToken> tokens;
            std::size_t              pos{0};

            [[nodiscard]] const XLToken& current() const;
            [[nodiscard]] const XLToken& peek(std::size_t offset = 1) const;
            const XLToken& consume();
            bool matchKind(XLTokenKind k);
        };

        static std::unique_ptr<XLASTNode> parseExpr(ParseContext& ctx, int minPrec = 0);
        static std::unique_ptr<XLASTNode> parseUnary(ParseContext& ctx);
        static std::unique_ptr<XLASTNode> parsePrimary(ParseContext& ctx);
        static std::unique_ptr<XLASTNode> parseFuncCall(std::string name, ParseContext& ctx);

        static int precedence(XLTokenKind k);
        static bool isRightAssoc(XLTokenKind k);
    };

    // =========================================================================
    // Evaluator / Engine
    // =========================================================================

    /**
     * @brief Callback type: resolves a cell reference string to a cell value.
     * @details The string is the raw reference text, e.g. "A1", "$B$2", "Sheet1!C3",
     *          or "A1:B10" for a range (the engine calls this once per distinct ref).
     *          Return `XLCellValue{}` (empty) for unknown or out-of-range cells.
     *
     *          **Range refs** are passed as-is.  The engine detects the colon in the
     *          text and will call the resolver with each individual cell in the range
     *          when it needs to expand the range.  The resolver only needs to handle
     *          single-cell refs; range expansion is done internally.
     */
    using XLCellResolver = std::function<XLCellValue(std::string_view ref)>;

    /**
     * @brief Lightweight formula evaluation engine.
     *
     * @details Usage:
     * @code
     *   XLFormulaEngine engine;
     *   auto resolver = XLFormulaEngine::makeResolver(worksheet);
     *   XLCellValue result = engine.evaluate("SUM(A1:C1)", resolver);
     * @endcode
     *
     * The engine is **thread-safe for concurrent evaluate() calls** after construction
     * (the function table is built once in the constructor and is read-only thereafter).
     */
    class OPENXLSX_EXPORT XLFormulaEngine
    {
    public:
        XLFormulaEngine();
        ~XLFormulaEngine() = default;

        XLFormulaEngine(const XLFormulaEngine&)            = delete;
        XLFormulaEngine& operator=(const XLFormulaEngine&) = delete;
        XLFormulaEngine(XLFormulaEngine&&)                 = default;
        XLFormulaEngine& operator=(XLFormulaEngine&&)      = default;

        /**
         * @brief Evaluate a formula string.
         * @param formula The formula text (with or without leading '=').
         * @param resolver Callback to look up cell values.  May be empty if the formula
         *        contains no cell references.
         * @return The computed XLCellValue; an error value on evaluation failure.
         */
        [[nodiscard]] XLCellValue evaluate(std::string_view formula,
                                           const XLCellResolver& resolver = {}) const;

        /**
         * @brief Create a CellResolver that reads live values from an XLWorksheet.
         * @param wks The source worksheet.
         * @return A resolver callable capturing a reference to @p wks.
         * @note The returned resolver is only valid while @p wks is alive.
         */
        [[nodiscard]] static XLCellResolver makeResolver(const XLWorksheet& wks);

    private:
        // ---- Internal evaluation helpers ----

        /**
         * @brief Expand a range string "A1:B3" using the resolver into a flat value list.
         */
        static std::vector<XLCellValue> expandRange(std::string_view rangeRef,
                                                     const XLCellResolver& resolver);

        /**
         * @brief Collect numeric values from a mixed argument list (scalars + range vectors).
         */
        static std::vector<double> collectNumbers(
            const std::vector<XLCellValue>& flat,
            bool countBlanks = false);

        /**
         * @brief Evaluate a single AST node recursively.
         * @throws XLFormulaError on unrecoverable error.
         */
        [[nodiscard]] XLCellValue evalNode(const XLASTNode& node,
                                           const XLCellResolver& resolver) const;

        /**
         * @brief Expand a function argument into a flat vector of XLCellValue.
         *        Handles both scalar values and ranges.
         */
        [[nodiscard]] std::vector<XLCellValue> expandArg(const XLASTNode& argNode,
                                                          const XLCellResolver& resolver) const;

        // ---- Built-in function table ----
        // Each entry maps an uppercase function name to its implementation.
        using FuncArgs  = std::vector<XLCellValue>;     ///< all arguments flattened
        using FuncImpl  = std::function<XLCellValue(const std::vector<std::vector<XLCellValue>>&)>;
        std::unordered_map<std::string, FuncImpl> m_functions;

        void registerBuiltins();

        // ---- Helpers registered as lambdas in registerBuiltins() ----
        static XLCellValue fnSum(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnAverage(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnMin(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnMax(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnCount(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnCounta(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnIf(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnAnd(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnOr(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnNot(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnIferror(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnAbs(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnRound(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnRoundup(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnRounddown(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnSqrt(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnInt(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnMod(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnPower(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnVlookup(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnHlookup(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnIndex(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnMatch(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnConcatenate(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnLen(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnLeft(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnRight(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnMid(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnUpper(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnLower(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnTrim(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnText(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnIsnumber(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnIsblank(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnIserror(const std::vector<std::vector<XLCellValue>>& args);
        static XLCellValue fnIstext(const std::vector<std::vector<XLCellValue>>& args);
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLFORMULAENGINE_HPP
