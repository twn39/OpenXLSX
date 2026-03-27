// ===== External Includes ===== //
#include <algorithm>
#include <cassert>
#include <cctype>
#include <cmath>
#include <fmt/format.h>
#include <functional>
#include <string>
#include <unordered_map>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLCellReference.hpp"
#include "XLException.hpp"
#include "XLFormulaEngine.hpp"
#include "XLWorksheet.hpp"

using namespace OpenXLSX;

// =============================================================================
// Internal helpers (anonymous namespace)
// =============================================================================
namespace
{
    // Convert a value to double; return NaN for non-numeric/empty.
    double toDouble(const XLCellValue& v)
    {
        switch (v.type()) {
            case XLValueType::Integer: return static_cast<double>(v.get<int64_t>());
            case XLValueType::Float:   return v.get<double>();
            case XLValueType::Boolean: return v.get<bool>() ? 1.0 : 0.0;
            default:                   return std::numeric_limits<double>::quiet_NaN();
        }
    }

    bool isNumeric(const XLCellValue& v)
    {
        return v.type() == XLValueType::Integer ||
               v.type() == XLValueType::Float   ||
               v.type() == XLValueType::Boolean;
    }

    bool isEmpty(const XLCellValue& v) { return v.type() == XLValueType::Empty; }

    bool isError(const XLCellValue& v) { return v.type() == XLValueType::Error; }

    std::string toString(const XLCellValue& v)
    {
        switch (v.type()) {
            case XLValueType::String:   return v.get<std::string>();
            case XLValueType::Integer:  return fmt::format("{}", v.get<int64_t>());
            case XLValueType::Float:    return fmt::format("{}", v.get<double>());
            case XLValueType::Boolean:  return v.get<bool>() ? "TRUE" : "FALSE";
            default:                    return "";
        }
    }

    XLCellValue errValue()  { XLCellValue r; r.setError("#VALUE!"); return r; }
    XLCellValue errDiv0()   { XLCellValue r; r.setError("#DIV/0!"); return r; }
    XLCellValue errNA()     { XLCellValue r; r.setError("#N/A");    return r; }
    XLCellValue errNum()    { XLCellValue r; r.setError("#NUM!");   return r; }
    XLCellValue errRef()    { XLCellValue r; r.setError("#REF!");   return r; }
    XLCellValue errName()   { XLCellValue r; r.setError("#NAME?");  return r; }

    // Trim leading/trailing whitespace
    std::string strTrim(std::string s)
    {
        const char* ws = " \t\r\n";
        s.erase(0, s.find_first_not_of(ws));
        s.erase(s.find_last_not_of(ws) + 1);
        return s;
    }

    // Flatten a vector-of-vectors into one vector
    std::vector<XLCellValue> flatten(const std::vector<std::vector<XLCellValue>>& args)
    {
        std::vector<XLCellValue> out;
        for (const auto& vec : args)
            for (const auto& v : vec)
                out.push_back(v);
        return out;
    }

    // Collect numeric values from a flat list
    std::vector<double> numerics(const std::vector<XLCellValue>& flat)
    {
        std::vector<double> out;
        for (const auto& v : flat)
            if (isNumeric(v))
                out.push_back(toDouble(v));
        return out;
    }
}    // namespace

// =============================================================================
// Lexer
// =============================================================================

std::vector<XLToken> XLFormulaLexer::tokenize(std::string_view formula)
{
    std::vector<XLToken> tokens;

    // Strip leading '='
    std::size_t i = 0;
    if (!formula.empty() && formula[i] == '=') ++i;

    const auto len = formula.size();

    auto emit = [&](XLTokenKind k, std::string text = {}, double num = 0.0, bool b = false) {
        XLToken t;
        t.kind = k; t.text = std::move(text); t.number = num; t.boolean = b;
        tokens.push_back(std::move(t));
    };

    while (i < len) {
        char c = formula[i];

        // --- Whitespace ---
        if (std::isspace(static_cast<unsigned char>(c))) { ++i; continue; }

        // --- String literal ---
        if (c == '"') {
            std::string s;
            ++i;
            while (i < len) {
                if (formula[i] == '"') {
                    ++i;
                    if (i < len && formula[i] == '"') { s += '"'; ++i; } // escaped quote
                    else break;
                } else {
                    s += formula[i++];
                }
            }
            emit(XLTokenKind::String, s);
            continue;
        }

        // --- Number ---
        if (std::isdigit(static_cast<unsigned char>(c)) || (c == '.' && i+1 < len && std::isdigit(static_cast<unsigned char>(formula[i+1])))) {
            std::string numStr;
            while (i < len && (std::isdigit(static_cast<unsigned char>(formula[i])) || formula[i] == '.')) numStr += formula[i++];
            // Exponent
            if (i < len && (formula[i] == 'e' || formula[i] == 'E')) {
                numStr += formula[i++];
                if (i < len && (formula[i] == '+' || formula[i] == '-')) numStr += formula[i++];
                while (i < len && std::isdigit(static_cast<unsigned char>(formula[i]))) numStr += formula[i++];
            }
            double val = 0.0;
            try { val = std::stod(numStr); } catch(...) {}
            emit(XLTokenKind::Number, numStr, val);
            continue;
        }

        // --- Identifier or cell ref or bool ---
        if (std::isalpha(static_cast<unsigned char>(c)) || c == '$' || c == '_') {
            std::string ident;
            // Collect potentially qualified name: letters, digits, $, _, !
            while (i < len && (std::isalnum(static_cast<unsigned char>(formula[i])) ||
                                formula[i] == '$' || formula[i] == '_' || formula[i] == '!'))
            {
                ident += formula[i++];
            }
            // Check for range colon right after reference
            if (i < len && formula[i] == ':') {
                // Might be a range; collect the second ref
                std::string second;
                ++i; // skip ':'
                while (i < len && (std::isalnum(static_cast<unsigned char>(formula[i])) || formula[i] == '$'))
                    second += formula[i++];
                if (!second.empty()) {
                    emit(XLTokenKind::CellRef, ident + ":" + second);
                } else {
                    emit(XLTokenKind::CellRef, ident);
                    emit(XLTokenKind::Colon);
                }
                continue;
            }

            // Boolean?
            std::string upper = ident;
            std::transform(upper.begin(), upper.end(), upper.begin(), ::toupper);
            if (upper == "TRUE")  { emit(XLTokenKind::Bool, ident, 1.0, true);  continue; }
            if (upper == "FALSE") { emit(XLTokenKind::Bool, ident, 0.0, false); continue; }

            // Cell reference? (letters then digits, optional $)
            // Pattern: optional$ + 1-3 letters + optional$ + 1-7 digits  (within ident already collected)
            {
                std::size_t j = 0;
                if (j < ident.size() && ident[j] == '$') ++j;
                std::size_t alphaStart = j;
                while (j < ident.size() && std::isalpha(static_cast<unsigned char>(ident[j]))) ++j;
                std::size_t alphaCount = j - alphaStart;
                if (j < ident.size() && ident[j] == '$') ++j;
                std::size_t digitStart = j;
                while (j < ident.size() && std::isdigit(static_cast<unsigned char>(ident[j]))) ++j;
                std::size_t digitCount = j - digitStart;
                if (alphaCount >= 1 && alphaCount <= 3 && digitCount >= 1 && j == ident.size()) {
                    emit(XLTokenKind::CellRef, ident); continue;
                }
                // Sheet-qualified ref like "Sheet1!A1"
                auto bang = ident.find('!');
                if (bang != std::string::npos) { emit(XLTokenKind::CellRef, ident); continue; }
            }

            emit(XLTokenKind::Ident, ident);
            continue;
        }

        // --- Single or double char operators ---
        switch (c) {
            case '+': emit(XLTokenKind::Plus);    ++i; break;
            case '-': emit(XLTokenKind::Minus);   ++i; break;
            case '*': emit(XLTokenKind::Star);    ++i; break;
            case '/': emit(XLTokenKind::Slash);   ++i; break;
            case '^': emit(XLTokenKind::Caret);   ++i; break;
            case '%': emit(XLTokenKind::Percent); ++i; break;
            case '&': emit(XLTokenKind::Amp);     ++i; break;
            case '(': emit(XLTokenKind::LParen);  ++i; break;
            case ')': emit(XLTokenKind::RParen);  ++i; break;
            case ',': emit(XLTokenKind::Comma);   ++i; break;
            case ';': emit(XLTokenKind::Semicolon); ++i; break;
            case ':': emit(XLTokenKind::Colon);   ++i; break;
            case '=': emit(XLTokenKind::Eq);      ++i; break;
            case '<':
                ++i;
                if (i < len && formula[i] == '=') { emit(XLTokenKind::Le); ++i; }
                else if (i < len && formula[i] == '>') { emit(XLTokenKind::NEq); ++i; }
                else emit(XLTokenKind::Lt);
                break;
            case '>':
                ++i;
                if (i < len && formula[i] == '=') { emit(XLTokenKind::Ge); ++i; }
                else emit(XLTokenKind::Gt);
                break;
            default:
                emit(XLTokenKind::Error, std::string(1, c)); ++i; break;
        }
    }

    emit(XLTokenKind::End);
    return tokens;
}

// =============================================================================
// Parser helpers
// =============================================================================

const XLToken& XLFormulaParser::ParseContext::current() const
{
    Expects(pos < tokens.size());
    return tokens[pos];
}

const XLToken& XLFormulaParser::ParseContext::peek(std::size_t offset) const
{
    auto idx = pos + offset;
    return idx < tokens.size() ? tokens[idx] : tokens[tokens.size() - 1]; // return End if past
}

const XLToken& XLFormulaParser::ParseContext::consume()
{
    const XLToken& t = current();
    if (t.kind != XLTokenKind::End) ++pos;
    return t;
}

bool XLFormulaParser::ParseContext::matchKind(XLTokenKind k)
{
    if (current().kind == k) { consume(); return true; }
    return false;
}

// Operator precedence (higher = binds tighter)
int XLFormulaParser::precedence(XLTokenKind k)
{
    switch (k) {
        case XLTokenKind::Amp:     return 1;   // string concat
        case XLTokenKind::Eq:
        case XLTokenKind::NEq:
        case XLTokenKind::Lt:
        case XLTokenKind::Le:
        case XLTokenKind::Gt:
        case XLTokenKind::Ge:      return 2;   // comparison
        case XLTokenKind::Plus:
        case XLTokenKind::Minus:   return 3;   // add/sub
        case XLTokenKind::Star:
        case XLTokenKind::Slash:   return 4;   // mul/div
        case XLTokenKind::Caret:   return 5;   // power (right-assoc)
        default:                   return -1;  // not a binary operator
    }
}

bool XLFormulaParser::isRightAssoc(XLTokenKind k)
{
    return k == XLTokenKind::Caret;
}

std::unique_ptr<XLASTNode> XLFormulaParser::parse(gsl::span<const XLToken> tokens)
{
    ParseContext ctx{ tokens, 0 };
    auto node = parseExpr(ctx, 0);
    return node;
}

// Pratt-style expression parser
std::unique_ptr<XLASTNode> XLFormulaParser::parseExpr(ParseContext& ctx, int minPrec)
{
    auto lhs = parseUnary(ctx);

    while (true) {
        const XLToken& op = ctx.current();
        int prec = precedence(op.kind);
        if (prec < minPrec) break;

        XLTokenKind opKind = op.kind;
        ctx.consume();

        int nextPrec = isRightAssoc(opKind) ? prec : prec + 1;
        auto rhs = parseExpr(ctx, nextPrec);

        auto binNode = std::make_unique<XLASTNode>(XLNodeKind::BinOp);
        binNode->op = opKind;
        binNode->children.push_back(std::move(lhs));
        binNode->children.push_back(std::move(rhs));
        lhs = std::move(binNode);
    }

    return lhs;
}

std::unique_ptr<XLASTNode> XLFormulaParser::parseUnary(ParseContext& ctx)
{
    // Unary minus
    if (ctx.current().kind == XLTokenKind::Minus) {
        ctx.consume();
        auto operand = parseUnary(ctx);
        auto node = std::make_unique<XLASTNode>(XLNodeKind::UnaryOp);
        node->op = XLTokenKind::Minus;
        node->children.push_back(std::move(operand));
        return node;
    }
    // Unary plus (no-op)
    if (ctx.current().kind == XLTokenKind::Plus) {
        ctx.consume();
        return parseUnary(ctx);
    }
    auto expr = parsePrimary(ctx);
    // Postfix percent
    if (ctx.current().kind == XLTokenKind::Percent) {
        ctx.consume();
        auto node = std::make_unique<XLASTNode>(XLNodeKind::UnaryOp);
        node->op = XLTokenKind::Percent;
        node->children.push_back(std::move(expr));
        return node;
    }
    return expr;
}

std::unique_ptr<XLASTNode> XLFormulaParser::parsePrimary(ParseContext& ctx)
{
    const XLToken& tok = ctx.current();

    if (tok.kind == XLTokenKind::Number) {
        ctx.consume();
        auto node = std::make_unique<XLASTNode>(XLNodeKind::Number);
        node->number = tok.number;
        node->text   = tok.text;
        return node;
    }

    if (tok.kind == XLTokenKind::String) {
        ctx.consume();
        auto node = std::make_unique<XLASTNode>(XLNodeKind::StringLit);
        node->text = tok.text;
        return node;
    }

    if (tok.kind == XLTokenKind::Bool) {
        ctx.consume();
        auto node = std::make_unique<XLASTNode>(XLNodeKind::BoolLit);
        node->boolean = tok.boolean;
        return node;
    }

    if (tok.kind == XLTokenKind::CellRef) {
        ctx.consume();
        bool isRange = (tok.text.find(':') != std::string::npos);
        auto node = std::make_unique<XLASTNode>(isRange ? XLNodeKind::Range : XLNodeKind::CellRef);
        node->text = tok.text;
        return node;
    }

    if (tok.kind == XLTokenKind::Ident) {
        std::string name = tok.text;
        ctx.consume();
        // Function call?
        if (ctx.current().kind == XLTokenKind::LParen) {
            return parseFuncCall(std::move(name), ctx);
        }
        // Named range / constant treated as cell ref (engine will try resolver)
        auto node = std::make_unique<XLASTNode>(XLNodeKind::CellRef);
        node->text = name;
        return node;
    }

    if (tok.kind == XLTokenKind::LParen) {
        ctx.consume();
        auto inner = parseExpr(ctx, 0);
        if (ctx.current().kind == XLTokenKind::RParen) ctx.consume();
        return inner;
    }

    // Error literal (#DIV/0! etc.) - just capture the token as error
    auto node = std::make_unique<XLASTNode>(XLNodeKind::ErrorLit);
    node->text = tok.text;
    ctx.consume();
    return node;
}

std::unique_ptr<XLASTNode> XLFormulaParser::parseFuncCall(std::string name, ParseContext& ctx)
{
    Expects(ctx.current().kind == XLTokenKind::LParen);
    ctx.consume(); // eat '('

    auto node = std::make_unique<XLASTNode>(XLNodeKind::FuncCall);
    // Store function name as uppercase
    std::transform(name.begin(), name.end(), name.begin(), ::toupper);
    node->text = std::move(name);

    // Parse argument list (comma or semicolon separated)
    while (ctx.current().kind != XLTokenKind::RParen &&
           ctx.current().kind != XLTokenKind::End)
    {
        node->children.push_back(parseExpr(ctx, 0));
        if (ctx.current().kind == XLTokenKind::Comma ||
            ctx.current().kind == XLTokenKind::Semicolon)
        {
            ctx.consume();
        } else {
            break;
        }
    }
    if (ctx.current().kind == XLTokenKind::RParen) ctx.consume();
    return node;
}

// =============================================================================
// Range expansion helper
// =============================================================================

std::vector<XLCellValue> XLFormulaEngine::expandRange(std::string_view rangeRef,
                                                        const XLCellResolver& resolver)
{
    std::vector<XLCellValue> result;
    if (!resolver) return result;

    auto colonPos = rangeRef.find(':');
    if (colonPos == std::string_view::npos) {
        // Single cell
        result.push_back(resolver(rangeRef));
        return result;
    }

    std::string startRef(rangeRef.substr(0, colonPos));
    std::string endRef  (rangeRef.substr(colonPos + 1));

    try {
        auto startCell = XLCellReference(startRef);
        auto endCell   = XLCellReference(endRef);
        uint32_t r1 = startCell.row(),    r2 = endCell.row();
        uint16_t c1 = startCell.column(), c2 = endCell.column();
        if (r1 > r2) std::swap(r1, r2);
        if (c1 > c2) std::swap(c1, c2);

        for (uint32_t r = r1; r <= r2; ++r)
            for (uint16_t c = c1; c <= c2; ++c)
                result.push_back(resolver(XLCellReference(r, c).address()));
    } catch (...) {
        XLCellValue err; err.setError("#REF!"); result.push_back(err);
    }
    return result;
}

// =============================================================================
// Evaluator – expandArg
// =============================================================================

std::vector<XLCellValue> XLFormulaEngine::expandArg(const XLASTNode& argNode,
                                                      const XLCellResolver& resolver) const
{
    if (argNode.kind == XLNodeKind::Range)
        return expandRange(argNode.text, resolver);

    // Evaluate normally and wrap in a single-element vector
    return { evalNode(argNode, resolver) };
}

// =============================================================================
// Evaluator – evalNode
// =============================================================================

XLCellValue XLFormulaEngine::evalNode(const XLASTNode& node, const XLCellResolver& resolver) const
{
    switch (node.kind) {
        case XLNodeKind::Number:    return XLCellValue(node.number);
        case XLNodeKind::StringLit: return XLCellValue(node.text);
        case XLNodeKind::BoolLit:   return XLCellValue(node.boolean);
        case XLNodeKind::ErrorLit:  { XLCellValue e; e.setError(node.text); return e; }

        case XLNodeKind::CellRef: {
            if (!resolver) return XLCellValue{};
            return resolver(node.text);
        }

        case XLNodeKind::Range: {
            // Range used as scalar = first cell value
            auto vals = expandRange(node.text, resolver);
            return vals.empty() ? XLCellValue{} : vals.front();
        }

        case XLNodeKind::UnaryOp: {
            Expects(node.children.size() == 1);
            auto val = evalNode(*node.children[0], resolver);
            if (node.op == XLTokenKind::Minus) {
                if (!isNumeric(val)) return errValue();
                double d = toDouble(val);
                if (val.type() == XLValueType::Integer)
                    return XLCellValue(static_cast<int64_t>(-d));
                return XLCellValue(-d);
            }
            if (node.op == XLTokenKind::Percent) {
                if (!isNumeric(val)) return errValue();
                return XLCellValue(toDouble(val) / 100.0);
            }
            return val;
        }

        case XLNodeKind::BinOp: {
            Expects(node.children.size() == 2);

            // String concat – evaluate early, no numeric coercion
            if (node.op == XLTokenKind::Amp) {
                auto lv = evalNode(*node.children[0], resolver);
                auto rv = evalNode(*node.children[1], resolver);
                if (isError(lv)) return lv;
                if (isError(rv)) return rv;
                return XLCellValue(toString(lv) + toString(rv));
            }

            auto lv = evalNode(*node.children[0], resolver);
            auto rv = evalNode(*node.children[1], resolver);
            if (isError(lv)) return lv;
            if (isError(rv)) return rv;

            // Arithmetic operators
            if (node.op == XLTokenKind::Plus  || node.op == XLTokenKind::Minus ||
                node.op == XLTokenKind::Star  || node.op == XLTokenKind::Slash ||
                node.op == XLTokenKind::Caret)
            {
                if (!isNumeric(lv) || !isNumeric(rv)) return errValue();
                double l = toDouble(lv), r = toDouble(rv);
                switch (node.op) {
                    case XLTokenKind::Plus:  return XLCellValue(l + r);
                    case XLTokenKind::Minus: return XLCellValue(l - r);
                    case XLTokenKind::Star:  return XLCellValue(l * r);
                    case XLTokenKind::Slash:
                        if (r == 0.0) return errDiv0();
                        return XLCellValue(l / r);
                    case XLTokenKind::Caret: return XLCellValue(std::pow(l, r));
                    default: break;
                }
            }

            // Comparison operators
            {
                bool result = false;
                // Numeric comparison
                if (isNumeric(lv) && isNumeric(rv)) {
                    double l = toDouble(lv), r = toDouble(rv);
                    switch (node.op) {
                        case XLTokenKind::Eq:  result = (l == r); break;
                        case XLTokenKind::NEq: result = (l != r); break;
                        case XLTokenKind::Lt:  result = (l < r);  break;
                        case XLTokenKind::Le:  result = (l <= r); break;
                        case XLTokenKind::Gt:  result = (l > r);  break;
                        case XLTokenKind::Ge:  result = (l >= r); break;
                        default: return errValue();
                    }
                } else {
                    // String comparison (case-insensitive like Excel)
                    std::string ls = toString(lv), rs = toString(rv);
                    std::transform(ls.begin(), ls.end(), ls.begin(), ::tolower);
                    std::transform(rs.begin(), rs.end(), rs.begin(), ::tolower);
                    switch (node.op) {
                        case XLTokenKind::Eq:  result = (ls == rs); break;
                        case XLTokenKind::NEq: result = (ls != rs); break;
                        case XLTokenKind::Lt:  result = (ls < rs);  break;
                        case XLTokenKind::Le:  result = (ls <= rs); break;
                        case XLTokenKind::Gt:  result = (ls > rs);  break;
                        case XLTokenKind::Ge:  result = (ls >= rs); break;
                        default: return errValue();
                    }
                }
                return XLCellValue(result);
            }
        }

        case XLNodeKind::FuncCall: {
            auto it = m_functions.find(node.text);
            if (it == m_functions.end()) return errName();

            // Build per-arg vectors (ranges are expanded, scalars wrapped)
            std::vector<std::vector<XLCellValue>> argVecs;
            argVecs.reserve(node.children.size());
            for (const auto& child : node.children)
                argVecs.push_back(expandArg(*child, resolver));

            try {
                return it->second(argVecs);
            } catch (const std::exception& ex) {
                XLCellValue e; e.setError(std::string("#ERROR: ") + ex.what()); return e;
            }
        }

        default:
            return errValue();
    }
}

// =============================================================================
// Evaluator – public evaluate()
// =============================================================================

XLCellValue XLFormulaEngine::evaluate(std::string_view formula,
                                       const XLCellResolver& resolver) const
{
    if (formula.empty()) return XLCellValue{};
    try {
        auto tokens = XLFormulaLexer::tokenize(formula);
        auto ast    = XLFormulaParser::parse(gsl::span<const XLToken>(tokens));
        return evalNode(*ast, resolver);
    } catch (const XLException&) {
        throw;
    } catch (const std::exception& ex) {
        XLCellValue e; e.setError(std::string("#ERROR: ") + ex.what()); return e;
    }
}

// =============================================================================
// makeResolver
// =============================================================================

XLCellResolver XLFormulaEngine::makeResolver(const XLWorksheet& wks)
{
    return [&wks](std::string_view ref) -> XLCellValue {
        try {
            // Strip sheet prefix (e.g. "Sheet1!A1" -> "A1")
            auto bang = ref.find('!');
            std::string cellAddr = std::string(bang != std::string_view::npos ? ref.substr(bang + 1) : ref);
            return wks.cell(cellAddr).value();
        } catch (...) {
            return XLCellValue{};
        }
    };
}

// =============================================================================
// Built-in function registrations
// =============================================================================

XLFormulaEngine::XLFormulaEngine() { registerBuiltins(); }

void XLFormulaEngine::registerBuiltins()
{
    m_functions["SUM"]         = fnSum;
    m_functions["AVERAGE"]     = fnAverage;
    m_functions["AVG"]         = fnAverage;   // alias
    m_functions["MIN"]         = fnMin;
    m_functions["MAX"]         = fnMax;
    m_functions["COUNT"]       = fnCount;
    m_functions["COUNTA"]      = fnCounta;
    m_functions["IF"]          = fnIf;
    m_functions["AND"]         = fnAnd;
    m_functions["OR"]          = fnOr;
    m_functions["NOT"]         = fnNot;
    m_functions["IFERROR"]     = fnIferror;
    m_functions["ABS"]         = fnAbs;
    m_functions["ROUND"]       = fnRound;
    m_functions["ROUNDUP"]     = fnRoundup;
    m_functions["ROUNDDOWN"]   = fnRounddown;
    m_functions["SQRT"]        = fnSqrt;
    m_functions["INT"]         = fnInt;
    m_functions["MOD"]         = fnMod;
    m_functions["POWER"]       = fnPower;
    m_functions["VLOOKUP"]     = fnVlookup;
    m_functions["HLOOKUP"]     = fnHlookup;
    m_functions["INDEX"]       = fnIndex;
    m_functions["MATCH"]       = fnMatch;
    m_functions["CONCATENATE"] = fnConcatenate;
    m_functions["CONCAT"]      = fnConcatenate; // alias
    m_functions["LEN"]         = fnLen;
    m_functions["LEFT"]        = fnLeft;
    m_functions["RIGHT"]       = fnRight;
    m_functions["MID"]         = fnMid;
    m_functions["UPPER"]       = fnUpper;
    m_functions["LOWER"]       = fnLower;
    m_functions["TRIM"]        = fnTrim;
    m_functions["TEXT"]        = fnText;
    m_functions["ISNUMBER"]    = fnIsnumber;
    m_functions["ISBLANK"]     = fnIsblank;
    m_functions["ISERROR"]     = fnIserror;
    m_functions["ISTEXT"]      = fnIstext;
}

// =============================================================================
// Built-in: Math / Statistical
// =============================================================================

XLCellValue XLFormulaEngine::fnSum(const std::vector<std::vector<XLCellValue>>& args)
{
    double total = 0.0;
    for (const auto& v : numerics(flatten(args))) total += v;
    return XLCellValue(total);
}

XLCellValue XLFormulaEngine::fnAverage(const std::vector<std::vector<XLCellValue>>& args)
{
    auto nums = numerics(flatten(args));
    if (nums.empty()) return errDiv0();
    double total = 0.0;
    for (double v : nums) total += v;
    return XLCellValue(total / static_cast<double>(nums.size()));
}

XLCellValue XLFormulaEngine::fnMin(const std::vector<std::vector<XLCellValue>>& args)
{
    auto nums = numerics(flatten(args));
    if (nums.empty()) return XLCellValue(0.0);
    return XLCellValue(*std::min_element(nums.begin(), nums.end()));
}

XLCellValue XLFormulaEngine::fnMax(const std::vector<std::vector<XLCellValue>>& args)
{
    auto nums = numerics(flatten(args));
    if (nums.empty()) return XLCellValue(0.0);
    return XLCellValue(*std::max_element(nums.begin(), nums.end()));
}

XLCellValue XLFormulaEngine::fnCount(const std::vector<std::vector<XLCellValue>>& args)
{
    int64_t cnt = 0;
    for (const auto& v : flatten(args))
        if (isNumeric(v)) ++cnt;
    return XLCellValue(cnt);
}

XLCellValue XLFormulaEngine::fnCounta(const std::vector<std::vector<XLCellValue>>& args)
{
    int64_t cnt = 0;
    for (const auto& v : flatten(args))
        if (!isEmpty(v)) ++cnt;
    return XLCellValue(cnt);
}

XLCellValue XLFormulaEngine::fnAbs(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    const auto& v = args[0][0];
    if (!isNumeric(v)) return errValue();
    return XLCellValue(std::abs(toDouble(v)));
}

XLCellValue XLFormulaEngine::fnSqrt(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double d = toDouble(args[0][0]);
    if (d < 0.0) return errNum();
    return XLCellValue(std::sqrt(d));
}

XLCellValue XLFormulaEngine::fnInt(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    if (!isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(std::floor(toDouble(args[0][0]))));
}

XLCellValue XLFormulaEngine::fnRound(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double num    = toDouble(args[0][0]);
    int    digits = static_cast<int>(toDouble(args[1][0]));
    double factor = std::pow(10.0, digits);
    return XLCellValue(std::round(num * factor) / factor);
}

XLCellValue XLFormulaEngine::fnRoundup(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double num    = toDouble(args[0][0]);
    int    digits = static_cast<int>(toDouble(args[1][0]));
    double factor = std::pow(10.0, digits);
    return XLCellValue(num >= 0 ? std::ceil(num * factor) / factor : std::floor(num * factor) / factor);
}

XLCellValue XLFormulaEngine::fnRounddown(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double num    = toDouble(args[0][0]);
    int    digits = static_cast<int>(toDouble(args[1][0]));
    double factor = std::pow(10.0, digits);
    return XLCellValue(num >= 0 ? std::floor(num * factor) / factor : std::ceil(num * factor) / factor);
}

XLCellValue XLFormulaEngine::fnMod(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double n = toDouble(args[0][0]), d = toDouble(args[1][0]);
    if (d == 0.0) return errDiv0();
    return XLCellValue(std::fmod(n, d));
}

XLCellValue XLFormulaEngine::fnPower(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    return XLCellValue(std::pow(toDouble(args[0][0]), toDouble(args[1][0])));
}

// =============================================================================
// Built-in: Logical
// =============================================================================

XLCellValue XLFormulaEngine::fnIf(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    const auto& cond = args[0][0];
    bool test = false;
    if (cond.type() == XLValueType::Boolean)      test = cond.get<bool>();
    else if (isNumeric(cond))                      test = (toDouble(cond) != 0.0);
    else if (cond.type() == XLValueType::String)   test = !cond.get<std::string>().empty();

    if (test)  return (args.size() > 1 && !args[1].empty()) ? args[1][0] : XLCellValue(true);
    else       return (args.size() > 2 && !args[2].empty()) ? args[2][0] : XLCellValue(false);
}

XLCellValue XLFormulaEngine::fnAnd(const std::vector<std::vector<XLCellValue>>& args)
{
    for (const auto& v : flatten(args)) {
        if (!isNumeric(v) && v.type() != XLValueType::Boolean) return errValue();
        if (!toDouble(v)) return XLCellValue(false);
    }
    return XLCellValue(true);
}

XLCellValue XLFormulaEngine::fnOr(const std::vector<std::vector<XLCellValue>>& args)
{
    for (const auto& v : flatten(args)) {
        if (!isNumeric(v) && v.type() != XLValueType::Boolean) return errValue();
        if (toDouble(v)) return XLCellValue(true);
    }
    return XLCellValue(false);
}

XLCellValue XLFormulaEngine::fnNot(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    const auto& v = args[0][0];
    if (!isNumeric(v) && v.type() != XLValueType::Boolean) return errValue();
    return XLCellValue(!static_cast<bool>(toDouble(v)));
}

XLCellValue XLFormulaEngine::fnIferror(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    return isError(args[0][0]) ? args[1][0] : args[0][0];
}

// =============================================================================
// Built-in: Lookup
// =============================================================================

XLCellValue XLFormulaEngine::fnVlookup(const std::vector<std::vector<XLCellValue>>& args)
{
    // VLOOKUP(lookup_value, table_array, col_index, [exact_match])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty())
        return errValue();

    const XLCellValue& lookupVal = args[0][0];
    const auto& table            = args[1];
    int colIdx = static_cast<int>(toDouble(args[2][0])); // 1-based
    bool exact = (args.size() < 4 || args[3].empty()) ? true
               : (toDouble(args[3][0]) == 0.0);

    if (colIdx < 1) return errValue();

    // Determine number of rows: table flattened row-major, but we need nCols.
    // We can read nCols from args[2] context but we don't have it here.
    // Convention: table_array is passed as a flat vector from range expansion;
    // we need to know nCols to stride. We'll get nCols from the engine's expandRange
    // call earlier. However, here we only see the flat vector.
    //
    // As a heuristic: if colIdx > total elements, error. Otherwise treat each
    // "row" as colIdx wide (since we don't know nCols without DOM access).
    // Store nCols in args[1] metadata is not available here.
    // Best approach: we know the table range text would have nCols but that info
    // was lost during expand. We approximate: stride = colIdx (search first col, return colIdx-th).
    // For proper VLOOKUP, caller must ensure table range has >= colIdx columns.

    int stride = colIdx; // minimum possible stride
    int nRows  = static_cast<int>(table.size()) / stride;
    if (nRows == 0 || static_cast<int>(table.size()) % stride != 0) {
        // Try to be more flexible – still search col 1 and return col colIdx
        // by guessing the actual number of columns from colIdx alone
        // (This is the best we can do without nCols metadata)
    }

    // Linear scan column 0 of each row
    for (int r = 0; r < nRows; ++r) {
        int cellIdx = r * stride;
        if (cellIdx >= static_cast<int>(table.size())) break;
        const XLCellValue& key = table[static_cast<std::size_t>(cellIdx)];

        bool match = false;
        if (exact) {
            if (isNumeric(key) && isNumeric(lookupVal))
                match = (toDouble(key) == toDouble(lookupVal));
            else
                match = (toString(key) == toString(lookupVal));
        } else {
            // Approximate: find largest key <= lookupVal
            if (isNumeric(key) && isNumeric(lookupVal))
                match = (toDouble(key) <= toDouble(lookupVal));
        }

        if (match) {
            int resultIdx = r * stride + (colIdx - 1);
            if (resultIdx < static_cast<int>(table.size()))
                return table[static_cast<std::size_t>(resultIdx)];
            return errRef();
        }
    }
    return errNA();
}

XLCellValue XLFormulaEngine::fnHlookup(const std::vector<std::vector<XLCellValue>>& args)
{
    // HLOOKUP(lookup_value, table_array, row_index, [exact_match])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty())
        return errValue();

    const XLCellValue& lookupVal = args[0][0];
    const auto& table            = args[1];
    int rowIdx = static_cast<int>(toDouble(args[2][0])); // 1-based row to return
    bool exact = (args.size() < 4 || args[3].empty()) ? true
               : (toDouble(args[3][0]) == 0.0);

    if (rowIdx < 1) return errValue();

    // HLOOKUP searches the first row. We don't know nCols here either.
    // Scan all elements in the first "row" (table[0..nRows-1] where stride = rowIdx).
    int nCols = static_cast<int>(table.size()) / rowIdx;
    for (int c = 0; c < nCols; ++c) {
        const XLCellValue& key = table[static_cast<std::size_t>(c)];
        bool match = false;
        if (exact) {
            if (isNumeric(key) && isNumeric(lookupVal))
                match = (toDouble(key) == toDouble(lookupVal));
            else
                match = (toString(key) == toString(lookupVal));
        } else {
            if (isNumeric(key) && isNumeric(lookupVal))
                match = (toDouble(key) <= toDouble(lookupVal));
        }
        if (match) {
            std::size_t resultIdx = static_cast<std::size_t>(c + (rowIdx - 1) * nCols);
            if (resultIdx < table.size()) return table[resultIdx];
            return errRef();
        }
    }
    return errNA();
}

XLCellValue XLFormulaEngine::fnIndex(const std::vector<std::vector<XLCellValue>>& args)
{
    // INDEX(array, row_num, [col_num])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    const auto& arr = args[0];
    int r = static_cast<int>(toDouble(args[1][0]));
    int c = (args.size() > 2 && !args[2].empty()) ? static_cast<int>(toDouble(args[2][0])) : 1;
    if (r < 1 || c < 1) return errValue();
    // We don't know nCols; assume 1D array or use c=1
    std::size_t idx = static_cast<std::size_t>(r - 1);
    if (c > 1) {
        // Can't determine nCols without metadata; best effort
        idx = static_cast<std::size_t>((r - 1) * c + (c - 1));
    }
    if (idx >= arr.size()) return errRef();
    return arr[idx];
}

XLCellValue XLFormulaEngine::fnMatch(const std::vector<std::vector<XLCellValue>>& args)
{
    // MATCH(lookup_value, lookup_array, [match_type])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    const XLCellValue& lookupVal = args[0][0];
    const auto& arr = args[1];
    int matchType = (args.size() > 2 && !args[2].empty()) ? static_cast<int>(toDouble(args[2][0])) : 1;

    if (matchType == 0) {
        // Exact match
        for (std::size_t i = 0; i < arr.size(); ++i) {
            bool match = false;
            if (isNumeric(arr[i]) && isNumeric(lookupVal))
                match = (toDouble(arr[i]) == toDouble(lookupVal));
            else
                match = (toString(arr[i]) == toString(lookupVal));
            if (match) return XLCellValue(static_cast<int64_t>(i + 1));
        }
        return errNA();
    }
    // Approximate match (matchType == 1: largest <= lookup)
    int64_t bestIdx = -1;
    double  bestVal = std::numeric_limits<double>::lowest();
    for (std::size_t i = 0; i < arr.size(); ++i) {
        if (isNumeric(arr[i]) && isNumeric(lookupVal)) {
            double d = toDouble(arr[i]);
            if (d <= toDouble(lookupVal) && d >= bestVal) { bestVal = d; bestIdx = static_cast<int64_t>(i + 1); }
        }
    }
    if (bestIdx < 0) return errNA();
    return XLCellValue(bestIdx);
}

// =============================================================================
// Built-in: Text
// =============================================================================

XLCellValue XLFormulaEngine::fnConcatenate(const std::vector<std::vector<XLCellValue>>& args)
{
    std::string result;
    for (const auto& v : flatten(args)) result += toString(v);
    return XLCellValue(result);
}

XLCellValue XLFormulaEngine::fnLen(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    return XLCellValue(static_cast<int64_t>(toString(args[0][0]).size()));
}

XLCellValue XLFormulaEngine::fnLeft(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    int64_t n = (args.size() > 1 && !args[1].empty()) ? static_cast<int64_t>(toDouble(args[1][0])) : 1;
    if (n < 0) return errValue();
    return XLCellValue(s.substr(0, static_cast<std::size_t>(n)));
}

XLCellValue XLFormulaEngine::fnRight(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    int64_t n = (args.size() > 1 && !args[1].empty()) ? static_cast<int64_t>(toDouble(args[1][0])) : 1;
    if (n < 0) return errValue();
    auto len = static_cast<int64_t>(s.size());
    auto start = len - n;
    if (start < 0) start = 0;
    return XLCellValue(s.substr(static_cast<std::size_t>(start)));
}

XLCellValue XLFormulaEngine::fnMid(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    std::string s = toString(args[0][0]);
    int64_t start = static_cast<int64_t>(toDouble(args[1][0])) - 1; // 1-based
    int64_t count = static_cast<int64_t>(toDouble(args[2][0]));
    if (start < 0 || count < 0) return errValue();
    return XLCellValue(s.substr(static_cast<std::size_t>(start),
                                 static_cast<std::size_t>(count)));
}

XLCellValue XLFormulaEngine::fnUpper(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    std::transform(s.begin(), s.end(), s.begin(), ::toupper);
    return XLCellValue(s);
}

XLCellValue XLFormulaEngine::fnLower(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    std::transform(s.begin(), s.end(), s.begin(), ::tolower);
    return XLCellValue(s);
}

XLCellValue XLFormulaEngine::fnTrim(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    return XLCellValue(strTrim(toString(args[0][0])));
}

XLCellValue XLFormulaEngine::fnText(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    // Simplified: just convert to string (format string ignored)
    return XLCellValue(toString(args[0][0]));
}

// =============================================================================
// Built-in: Info
// =============================================================================

XLCellValue XLFormulaEngine::fnIsnumber(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    return XLCellValue(isNumeric(args[0][0]));
}

XLCellValue XLFormulaEngine::fnIsblank(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(true);
    return XLCellValue(isEmpty(args[0][0]));
}

XLCellValue XLFormulaEngine::fnIserror(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    return XLCellValue(isError(args[0][0]));
}

XLCellValue XLFormulaEngine::fnIstext(const std::vector<std::vector<XLCellValue>>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    return XLCellValue(args[0][0].type() == XLValueType::String);
}
