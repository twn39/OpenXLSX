#include "XLFormulaUtils.hpp"
#include <fmt/format.h>
#include <limits>
#include <algorithm>

namespace OpenXLSX {
    // Convert a value to double; return NaN for non-numeric/empty.
    double toDouble(const XLCellValue& v)
    {
        switch (v.type()) {
            case XLValueType::Integer:
                return static_cast<double>(v.get<int64_t>());
            case XLValueType::Float:
                return v.get<double>();
            case XLValueType::Boolean:
                return v.get<bool>() ? 1.0 : 0.0;
            default:
                return std::numeric_limits<double>::quiet_NaN();
        }
    }

    bool isNumeric(const XLCellValue& v)
    { return v.type() == XLValueType::Integer || v.type() == XLValueType::Float || v.type() == XLValueType::Boolean; }

    bool isEmpty(const XLCellValue& v) { return v.type() == XLValueType::Empty; }

    bool isError(const XLCellValue& v) { return v.type() == XLValueType::Error; }

    std::string toString(const XLCellValue& v)
    {
        switch (v.type()) {
            case XLValueType::String:
                return v.get<std::string>();
            case XLValueType::Integer:
                return fmt::format("{}", v.get<int64_t>());
            case XLValueType::Float:
                return fmt::format("{}", v.get<double>());
            case XLValueType::Boolean:
                return v.get<bool>() ? "TRUE" : "FALSE";
            default:
                return "";
        }
    }

    XLCellValue errValue()
    {
        XLCellValue r;
        r.setError("#VALUE!");
        return r;
    }
    XLCellValue errDiv0()
    {
        XLCellValue r;
        r.setError("#DIV/0!");
        return r;
    }
    XLCellValue errNA()
    {
        XLCellValue r;
        r.setError("#N/A");
        return r;
    }
    XLCellValue errNum()
    {
        XLCellValue r;
        r.setError("#NUM!");
        return r;
    }
    XLCellValue errRef()
    {
        XLCellValue r;
        r.setError("#REF!");
        return r;
    }
    XLCellValue errName()
    {
        XLCellValue r;
        r.setError("#NAME?");
        return r;
    }

    // Trim leading/trailing whitespace
    std::string strTrim(std::string s)
    {
        const char* ws = " \t\r\n";
        s.erase(0, s.find_first_not_of(ws));
        s.erase(s.find_last_not_of(ws) + 1);
        return s;
    }

    // Collect numeric values from a list of arguments directly
    std::vector<double> numerics(const std::vector<XLFormulaArg>& args)
    {
        std::vector<double> out;
        for (const auto& arg : args)
            for (const auto& v : arg)
                if (isNumeric(v)) out.push_back(toDouble(v));
        return out;
    }

    // Collect numeric values from a single argument
    std::vector<double> numerics(const XLFormulaArg& arg)
    {
        std::vector<double> out;
        for (const auto& v : arg)
            if (isNumeric(v)) out.push_back(toDouble(v));
        return out;
    }
}
