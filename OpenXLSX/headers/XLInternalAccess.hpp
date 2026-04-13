#ifndef OPENXLSX_XLINTERNALACCESS_HPP
#define OPENXLSX_XLINTERNALACCESS_HPP

#include "OpenXLSX-Exports.hpp"

namespace OpenXLSX {
    class XLXmlFile;
    class XLWorkbook;
    class XLSheet;
    class XLWorksheet;
    class XLXmlData;
    class XLTableCollection;
    class XLPivotTable;
    class XLFormulaProxy;
    class XLStreamReader;
    class XLCell;
    class XLCellValueProxy;

    /**
     * @brief Passkey Idiom for internal access control.
     * Only the specified friend classes can instantiate this passkey.
     * This prevents public API users from calling internal XLDocument methods.
     */
    class OPENXLSX_EXPORT XLInternalAccess {
        friend class XLXmlFile;
        friend class XLWorkbook;
        friend class XLSheet;
        friend class XLWorksheet;
        friend class XLXmlData;
        friend class XLTableCollection;
        friend class XLPivotTable;
        friend class XLFormulaProxy;
        friend class XLStreamReader;
        friend class XLCell;
        friend class XLCellValueProxy;
        
        XLInternalAccess() = default; 
    };
}

#endif // OPENXLSX_XLINTERNALACCESS_HPP