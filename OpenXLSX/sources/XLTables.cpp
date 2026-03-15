// ===== External Includes ===== //
#include <pugixml.hpp>
#include <string_view>

// ===== OpenXLSX Includes ===== //
#include "XLTables.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

// ========== XLTables Member Functions

/**
 * @details The constructor creates an instance of the superclass, XLXmlFile
 */
XLTables::XLTables(XLXmlData* xmlData) : XLXmlFile(xmlData)
{
    if (xmlData && xmlData->getXmlType() != XLContentType::Table) throw XLInternalError("XLTables constructor: Invalid XML data.");
    if (xmlDocument().document_element().empty()) {
        xmlDocument().append_child("table");
    }
}

/**
 * @details Get the name of the table
 */
std::string XLTables::name() const
{
    return xmlDocument().document_element().attribute("name").value();
}

/**
 * @details Set the name of the table. Table names cannot have spaces.
 */
void XLTables::setName(std::string_view name)
{
    if (name.find(' ') != std::string_view::npos) {
        throw XLInputError("Table names cannot contain spaces.");
    }
    auto docNode = xmlDocument().document_element();
    appendAndSetAttribute(docNode, "name", std::string(name));
}

/**
 * @details Get the display name of the table
 */
std::string XLTables::displayName() const
{
    return xmlDocument().document_element().attribute("displayName").value();
}

/**
 * @details Set the display name of the table
 */
void XLTables::setDisplayName(std::string_view name)
{
    auto docNode = xmlDocument().document_element();
    appendAndSetAttribute(docNode, "displayName", std::string(name));
}

/**
 * @details Get the range reference of the table
 */
std::string XLTables::rangeReference() const
{
    return xmlDocument().document_element().attribute("ref").value();
}

/**
 * @details Set the range reference of the table
 */
void XLTables::setRangeReference(std::string_view ref)
{
    auto docNode = xmlDocument().document_element();
    appendAndSetAttribute(docNode, "ref", std::string(ref));
    
    // Also update autoFilter if it exists
    auto autoFilter = docNode.child("autoFilter");
    if (!autoFilter.empty()) {
        appendAndSetAttribute(autoFilter, "ref", std::string(ref));
    }
}

/**
 * @details Get the table style name
 */
std::string XLTables::styleName() const
{
    return xmlDocument().document_element().child("tableStyleInfo").attribute("name").value();
}

/**
 * @details Set the table style name
 */
void XLTables::setStyleName(std::string_view styleName)
{
    auto docNode = xmlDocument().document_element();
    auto info = docNode.child("tableStyleInfo");
    if (info.empty()) {
        info = docNode.append_child("tableStyleInfo");
    }
    appendAndSetAttribute(info, "name", std::string(styleName));
}

/**
 * @details Check if row stripes are shown
 */
bool XLTables::showRowStripes() const
{
    return xmlDocument().document_element().child("tableStyleInfo").attribute("showRowStripes").as_bool();
}

/**
 * @details Set whether row stripes are shown
 */
void XLTables::setShowRowStripes(bool show)
{
    auto info = xmlDocument().document_element().child("tableStyleInfo");
    if (info.empty()) info = xmlDocument().document_element().append_child("tableStyleInfo");
    appendAndSetAttribute(info, "showRowStripes", show ? "1" : "0");
}

/**
 * @details Check if column stripes are shown
 */
bool XLTables::showColumnStripes() const
{
    return xmlDocument().document_element().child("tableStyleInfo").attribute("showColumnStripes").as_bool();
}

/**
 * @details Set whether column stripes are shown
 */
void XLTables::setShowColumnStripes(bool show)
{
    auto info = xmlDocument().document_element().child("tableStyleInfo");
    if (info.empty()) info = xmlDocument().document_element().append_child("tableStyleInfo");
    appendAndSetAttribute(info, "showColumnStripes", show ? "1" : "0");
}

/**
 * @details Check if the first column is highlighted
 */
bool XLTables::showFirstColumn() const
{
    return xmlDocument().document_element().child("tableStyleInfo").attribute("showFirstColumn").as_bool();
}

/**
 * @details Set whether the first column is highlighted
 */
void XLTables::setShowFirstColumn(bool show)
{
    auto info = xmlDocument().document_element().child("tableStyleInfo");
    if (info.empty()) info = xmlDocument().document_element().append_child("tableStyleInfo");
    appendAndSetAttribute(info, "showFirstColumn", show ? "1" : "0");
}

/**
 * @details Check if the last column is highlighted
 */
bool XLTables::showLastColumn() const
{
    return xmlDocument().document_element().child("tableStyleInfo").attribute("showLastColumn").as_bool();
}

/**
 * @details Set whether the last column is highlighted
 */
void XLTables::setShowLastColumn(bool show)
{
    auto info = xmlDocument().document_element().child("tableStyleInfo");
    if (info.empty()) info = xmlDocument().document_element().append_child("tableStyleInfo");
    appendAndSetAttribute(info, "showLastColumn", show ? "1" : "0");
}

/**
 * @details Print the underlying XML using pugixml::xml_node::print
 */
void XLTables::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print(ostr); }
