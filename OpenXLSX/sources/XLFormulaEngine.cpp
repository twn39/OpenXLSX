// ===== External Includes ===== //
#include <algorithm>
#include <cassert>
#include <cctype>
#include <cmath>
#include <ctime>
#include <fmt/format.h>
#include <functional>
#include <numeric>
#include <string>
#include <unordered_map>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLCellReference.hpp"
#include "XLDateTime.hpp"
#include "XLException.hpp"
#include "XLFormulaEngine.hpp"
#include "XLWorksheet.hpp"

using namespace OpenXLSX;

// =============================================================================
#include "XLFormulaUtils.hpp"

// =============================================================================
// Lexer
// =============================================================================

XLFormulaArg XLFormulaEngine::expandArg(const XLASTNode& argNode, const XLCellResolver& resolver) const
{
    if (argNode.kind == XLNodeKind::Range) return expandRange(argNode.text, resolver);

    // Evaluate normally and wrap in a single-element scalar
    return XLFormulaArg(evalNode(argNode, resolver));
}

// =============================================================================
// Evaluator – evalNode
// =============================================================================

XLCellValue XLFormulaEngine::evalNode(const XLASTNode& node, const XLCellResolver& resolver) const
{
    switch (node.kind) {
        case XLNodeKind::Number:
            return XLCellValue(node.number);
        case XLNodeKind::StringLit:
            return XLCellValue(node.text);
        case XLNodeKind::BoolLit:
            return XLCellValue(node.boolean);
        case XLNodeKind::ErrorLit: {
            XLCellValue e;
            e.setError(node.text);
            return e;
        }

        case XLNodeKind::CellRef: {
            if (!resolver) return XLCellValue{};
            return resolver(node.text);
        }

        case XLNodeKind::Range: {
            // Range used as scalar = first cell value
            auto vals = expandRange(node.text, resolver);
            return vals.empty() ? XLCellValue{} : vals[0];
        }

        case XLNodeKind::UnaryOp: {
            Expects(node.children.size() == 1);
            auto val = evalNode(*node.children[0], resolver);
            if (node.op == XLTokenKind::Minus) {
                if (!isNumeric(val)) return errValue();
                double d = toDouble(val);
                if (val.type() == XLValueType::Integer) return XLCellValue(static_cast<int64_t>(-d));
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
            if (node.op == XLTokenKind::Plus || node.op == XLTokenKind::Minus || node.op == XLTokenKind::Star ||
                node.op == XLTokenKind::Slash || node.op == XLTokenKind::Caret)
            {
                if (!isNumeric(lv) || !isNumeric(rv)) return errValue();
                double l = toDouble(lv), r = toDouble(rv);
                switch (node.op) {
                    case XLTokenKind::Plus:
                        return XLCellValue(l + r);
                    case XLTokenKind::Minus:
                        return XLCellValue(l - r);
                    case XLTokenKind::Star:
                        return XLCellValue(l * r);
                    case XLTokenKind::Slash:
                        if (r == 0.0) return errDiv0();
                        return XLCellValue(l / r);
                    case XLTokenKind::Caret:
                        return XLCellValue(std::pow(l, r));
                    default:
                        break;
                }
            }

            // Comparison operators
            {
                bool result = false;
                // Numeric comparison
                if (isNumeric(lv) && isNumeric(rv)) {
                    double l = toDouble(lv), r = toDouble(rv);
                    switch (node.op) {
                        case XLTokenKind::Eq:
                            result = (l == r);
                            break;
                        case XLTokenKind::NEq:
                            result = (l != r);
                            break;
                        case XLTokenKind::Lt:
                            result = (l < r);
                            break;
                        case XLTokenKind::Le:
                            result = (l <= r);
                            break;
                        case XLTokenKind::Gt:
                            result = (l > r);
                            break;
                        case XLTokenKind::Ge:
                            result = (l >= r);
                            break;
                        default:
                            return errValue();
                    }
                }
                else {
                    // String comparison (case-insensitive like Excel)
                    std::string ls = toString(lv), rs = toString(rv);
                    std::transform(ls.begin(), ls.end(), ls.begin(), ::tolower);
                    std::transform(rs.begin(), rs.end(), rs.begin(), ::tolower);
                    switch (node.op) {
                        case XLTokenKind::Eq:
                            result = (ls == rs);
                            break;
                        case XLTokenKind::NEq:
                            result = (ls != rs);
                            break;
                        case XLTokenKind::Lt:
                            result = (ls < rs);
                            break;
                        case XLTokenKind::Le:
                            result = (ls <= rs);
                            break;
                        case XLTokenKind::Gt:
                            result = (ls > rs);
                            break;
                        case XLTokenKind::Ge:
                            result = (ls >= rs);
                            break;
                        default:
                            return errValue();
                    }
                }
                return XLCellValue(result);
            }
        }

        case XLNodeKind::FuncCall: {
            auto it = m_functions.find(node.text);
            if (it == m_functions.end()) return errName();

            // Build per-arg vectors (ranges are expanded, scalars wrapped)
            std::vector<XLFormulaArg> argVecs;
            argVecs.reserve(node.children.size());
            for (const auto& child : node.children) argVecs.push_back(expandArg(*child, resolver));

            try {
                return it->second(argVecs);
            }
            catch (const std::exception& ex) {
                XLCellValue e;
                e.setError(std::string("#ERROR: ") + ex.what());
                return e;
            }
        }

        default:
            return errValue();
    }
}

// =============================================================================
// Evaluator – public evaluate()
// =============================================================================

XLCellValue XLFormulaEngine::evaluate(std::string_view formula, const XLCellResolver& resolver) const
{
    if (formula.empty()) return XLCellValue{};
    try {
        auto tokens = XLFormulaLexer::tokenize(formula);
        auto ast    = XLFormulaParser::parse(gsl::span<const XLToken>(tokens));
        return evalNode(*ast, resolver);
    }
    catch (const XLException&) {
        throw;
    }
    catch (const std::exception& ex) {
        XLCellValue e;
        e.setError(std::string("#ERROR: ") + ex.what());
        return e;
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
            auto        bang     = ref.find('!');
            std::string cellAddr = std::string(bang != std::string_view::npos ? ref.substr(bang + 1) : ref);
            return wks.cell(cellAddr).value();
        }
        catch (...) {
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
    m_functions["AVG"]         = fnAverage;    // alias
    m_functions["MIN"]         = fnMin;
    m_functions["MAX"]         = fnMax;
    m_functions["COUNT"]       = fnCount;
    m_functions["COUNTA"]      = fnCounta;
    m_functions["IF"]          = fnIf;
    m_functions["IFS"]         = fnIfs;
    m_functions["SWITCH"]      = fnSwitch;
    m_functions["AND"]         = fnAnd;
    m_functions["OR"]          = fnOr;
    m_functions["NOT"]         = fnNot;
    m_functions["IFERROR"]     = fnIferror;
    m_functions["ABS"]         = fnAbs;
    m_functions["ROUND"]       = fnRound;
    m_functions["ROUNDUP"]     = fnRoundup;
    m_functions["ROUNDDOWN"]   = fnRounddown;
    m_functions["SQRT"]        = fnSqrt;
    m_functions["PI"]          = fnPi;
    m_functions["SIN"]         = fnSin;
    m_functions["COS"]         = fnCos;
    m_functions["TAN"]         = fnTan;
    m_functions["ASIN"]        = fnAsin;
    m_functions["ACOS"]        = fnAcos;
    m_functions["DEGREES"]     = fnDegrees;
    m_functions["RADIANS"]     = fnRadians;
    m_functions["RAND"]        = fnRand;
    m_functions["RANDBETWEEN"] = fnRandbetween;
    m_functions["INT"]         = fnInt;
    m_functions["MOD"]         = fnMod;
    m_functions["POWER"]       = fnPower;
    m_functions["VLOOKUP"]     = fnVlookup;
    m_functions["HLOOKUP"]     = fnHlookup;
    m_functions["XLOOKUP"]     = fnXlookup;
    m_functions["INDEX"]       = fnIndex;
    m_functions["MATCH"]       = fnMatch;
    m_functions["CONCATENATE"] = fnConcatenate;
    m_functions["CONCAT"]      = fnConcatenate;    // alias
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

    // ---- Date / Time ----
    m_functions["TODAY"]       = fnToday;
    m_functions["NOW"]         = fnNow;
    m_functions["DATE"]        = fnDate;
    m_functions["TIME"]        = fnTime;
    m_functions["YEAR"]        = fnYear;
    m_functions["MONTH"]       = fnMonth;
    m_functions["DAY"]         = fnDay;
    m_functions["HOUR"]        = fnHour;
    m_functions["MINUTE"]      = fnMinute;
    m_functions["SECOND"]      = fnSecond;
    m_functions["DAYS"]         = fnDays;
    m_functions["_xlfn.DAYS"]   = fnDays;
    m_functions["WEEKDAY"]      = fnWeekday;
    m_functions["EDATE"]        = fnEdate;
    m_functions["EOMONTH"]      = fnEomonth;
    m_functions["WORKDAY"]      = fnWorkday;
    m_functions["NETWORKDAYS"]  = fnNetworkdays;

    // ---- Financial ----
    m_functions["PMT"] = fnPmt;
    m_functions["FV"]  = fnFv;
    m_functions["PV"]  = fnPv;
    m_functions["NPV"] = fnNpv;

    // ---- Math extended ----
    m_functions["SUMPRODUCT"] = fnSumproduct;
    m_functions["CEILING"]    = fnCeil;
    m_functions["CEIL"]       = fnCeil;
    m_functions["FLOOR"]      = fnFloor;
    m_functions["LOG"]        = fnLog;
    m_functions["LOG10"]      = fnLog10;
    m_functions["EXP"]        = fnExp;
    m_functions["SIGN"]       = fnSign;

    // ---- Text extended ----
    m_functions["FIND"]       = fnFind;
    m_functions["SEARCH"]     = fnSearch;
    m_functions["SUBSTITUTE"] = fnSubstitute;
    m_functions["REPLACE"]    = fnReplace;
    m_functions["REPT"]       = fnRept;
    m_functions["EXACT"]      = fnExact;
    m_functions["T"]          = fnT;
    m_functions["VALUE"]      = fnValue;
    m_functions["TEXTJOIN"]   = fnTextjoin;
    m_functions["_xlfn.TEXTJOIN"] = fnTextjoin;
    m_functions["CLEAN"]      = fnClean;
    m_functions["PROPER"]     = fnProper;

    // ---- Statistical / Conditional ----
    m_functions["SUMIF"]        = fnSumif;
    m_functions["COUNTIF"]      = fnCountif;
    m_functions["SUMIFS"]       = fnSumifs;
    m_functions["COUNTIFS"]     = fnCountifs;
    m_functions["MAXIFS"]       = fnMaxifs;
    m_functions["_xlfn.MAXIFS"] = fnMaxifs;
    m_functions["MINIFS"]       = fnMinifs;
    m_functions["_xlfn.MINIFS"] = fnMinifs;
    m_functions["AVERAGEIF"]    = fnAverageif;
    m_functions["RANK"]       = fnRank;
    m_functions["RANK.EQ"]    = fnRank;
    m_functions["LARGE"]      = fnLarge;
    m_functions["SMALL"]      = fnSmall;
    m_functions["STDEV"]      = fnStdev;
    m_functions["STDEV.S"]    = fnStdev;
    m_functions["VAR"]        = fnVar;
    m_functions["VAR.S"]      = fnVar;
    m_functions["MEDIAN"]     = fnMedian;
    m_functions["COUNTBLANK"] = fnCountblank;

    // ---- Info extended ----
    m_functions["ISNA"]      = fnIsna;
    m_functions["IFNA"]      = fnIfna;
    m_functions["ISLOGICAL"] = fnIslogical;
    m_functions["ISNONTEXT"] = fnIsnontext;

    // ---- Easy Additions ----
    m_functions["TRUE"]            = fnTrue;
    m_functions["FALSE"]           = fnFalse;
    m_functions["ISEVEN"]          = fnIseven;
    m_functions["ISODD"]           = fnIsodd;
    m_functions["MROUND"]          = fnMround;
    m_functions["CEILING.MATH"]    = fnCeilingMath;
    m_functions["_xlfn.CEILING.MATH"] = fnCeilingMath;
    m_functions["FLOOR.MATH"]      = fnFloorMath;
    m_functions["_xlfn.FLOOR.MATH"]= fnFloorMath;
    m_functions["VAR.P"]           = fnVarp;
    m_functions["_xlfn.VAR.P"]     = fnVarp;
    m_functions["VARP"]            = fnVarp;
    m_functions["STDEV.P"]         = fnStdevp;
    m_functions["_xlfn.STDEV.P"]   = fnStdevp;
    m_functions["STDEVP"]          = fnStdevp;
    m_functions["VARA"]            = fnVara;
    m_functions["VARPA"]           = fnVarpa;
    m_functions["STDEVA"]          = fnStdeva;
    m_functions["STDEVPA"]         = fnStdevpa;
    m_functions["PERMUT"]          = fnPermut;
    m_functions["PERMUTATIONA"]    = fnPermutationa;
    m_functions["_xlfn.PERMUTATIONA"] = fnPermutationa;
    m_functions["FISHER"]          = fnFisher;
    m_functions["FISHERINV"]       = fnFisherinv;
    m_functions["STANDARDIZE"]     = fnStandardize;
    m_functions["PEARSON"]         = fnPearson;
    m_functions["CORREL"]          = fnPearson;
    m_functions["COVAR"]           = fnCovarianceP;
    m_functions["COVARIANCE.P"]    = fnCovarianceP;
    m_functions["COVARIANCE.S"]    = fnCovarianceS;
    m_functions["PERCENTILE"]      = fnPercentileInc;
    m_functions["PERCENTILE.INC"]  = fnPercentileInc;
    m_functions["PERCENTILE.EXC"]  = fnPercentileExc;
    m_functions["QUARTILE"]        = fnQuartileInc;
    m_functions["QUARTILE.INC"]    = fnQuartileInc;
    m_functions["QUARTILE.EXC"]    = fnQuartileExc;
    m_functions["TRIMMEAN"]        = fnTrimmean;
    m_functions["SLOPE"]           = fnSlope;
    m_functions["INTERCEPT"]       = fnIntercept;
    m_functions["RSQ"]             = fnRsq;
    m_functions["AVERAGEIFS"]      = fnAverageifs;
    m_functions["ISOWEEKNUM"]      = fnIsoweeknum;
    m_functions["_xlfn.ISOWEEKNUM"]= fnIsoweeknum;
    m_functions["WEEKNUM"]         = fnWeeknum;
    m_functions["DAYS360"]         = fnDays360;
    m_functions["NPER"]            = fnNper;
    m_functions["DB"]              = fnDb;
    m_functions["DDB"]             = fnDdb;

    // Fix mapped functions from earlier additions
    m_functions["ISERR"]           = fnIserr;
    m_functions["TRUNC"]           = fnTrunc;
    m_functions["SUMSQ"]           = fnSumsq;
    m_functions["SUMX2MY2"]        = fnSumx2my2;
    m_functions["SUMX2PY2"]        = fnSumx2py2;
    m_functions["SUMXMY2"]         = fnSumxmy2;
    m_functions["AVEDEV"]          = fnAvedev;
    m_functions["DEVSQ"]           = fnDevsq;
    m_functions["AVERAGEA"]        = fnAveragea;

    // Missing checklist functions
    m_functions["SLN"]             = fnSln;
    m_functions["SYD"]             = fnSyd;
    m_functions["CHAR"]            = fnChar;
    m_functions["UNICHAR"]         = fnUnichar;
    m_functions["CODE"]            = fnCode;
    m_functions["UNICODE"]         = fnUnicode;
    m_functions["NOW"]             = fnNow;
    m_functions["TODAY"]           = fnToday;
    m_functions["TRUE"]            = fnTrue;
    m_functions["FALSE"]           = fnFalse;
}

// =============================================================================
// Built-in: Math / Statistical
// =============================================================================

