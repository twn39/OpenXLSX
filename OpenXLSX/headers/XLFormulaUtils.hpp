#ifndef OPENXLSX_XLFORMULAUTILS_HPP
#define OPENXLSX_XLFORMULAUTILS_HPP

#include "XLCellValue.hpp"
#include "XLFormulaEngine.hpp"
#include <string>
#include <vector>

namespace OpenXLSX {
    double toDouble(const XLCellValue& v);
    bool isNumeric(const XLCellValue& v);
    bool isEmpty(const XLCellValue& v);
    bool isError(const XLCellValue& v);
    std::string toString(const XLCellValue& v);

    XLCellValue errValue();
    XLCellValue errDiv0();
    XLCellValue errNA();
    XLCellValue errNum();
    XLCellValue errRef();
    XLCellValue errName();

    std::string strTrim(std::string s);
    std::vector<double> numerics(const std::vector<XLFormulaArg>& args);
    std::vector<double> numerics(const XLFormulaArg& arg);
}

#endif
