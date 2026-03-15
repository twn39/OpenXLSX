// ===== External Includes ===== //
// #include <algorithm>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
// #include "XLCellRange.hpp"
// #include "XLDocument.hpp"
#include "XLTables.hpp"
#include "XLUtilities.hpp"    // OpenXLSX::ignore

using namespace OpenXLSX;

namespace OpenXLSX
{
    // placeholder for utility functions

}    // namespace OpenXLSX

// ========== XLTables Member Functions

/**
 * @details The constructor creates an instance of the superclass, XLXmlFile
 */
XLTables::XLTables(XLXmlData* xmlData) : XLXmlFile(xmlData)
{
    if (xmlData->getXmlType() != XLContentType::Table) throw XLInternalError("XLTables constructor: Invalid XML data.");
}

/**
 * @details getters
 */
/*
std::string XLTables::get(std::string cellRef) const
{
    OpenXLSX::ignore(cellRef);
    return "";
}
*/

/**
 * @details setters
 */
/*
bool XLTables::set(std::string cellRef)
{
    OpenXLSX::ignore(cellRef);
    return false;
}
*/

/**
 * @details Print the underlying XML using pugixml::xml_node::print
 */
void XLTables::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print(ostr); }
