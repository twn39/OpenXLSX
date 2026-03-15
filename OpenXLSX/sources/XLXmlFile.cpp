// ===== External Includes ===== //
#include <gsl/gsl>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLXmlFile.hpp"

using namespace OpenXLSX;

/**
 * @details The constructor creates a new object with the parent XLDocument and the file path as input, with
 * an optional input being a std::string with the XML data. If the XML data is provided by a string, any file with
 * the same path in the .zip file will be overwritten upon saving of the document. If no xmlData is provided,
 * the data will be read from the .zip file, using the given path.
 */
XLXmlFile::XLXmlFile(XLXmlData* xmlData) : m_xmlData(xmlData) {}

/**
 * @details This method sets the XML data with a std::string as input. The underlying XMLDocument reads the data.
 * When envoking the load_string method in PugiXML, the flag 'parse_ws_pcdata' is passed along with the default flags.
 * This will enable parsing of whitespace characters. If not set, Excel cells with only spaces will be returned as
 * empty strings, which is not what we want. The downside is that whitespace characters such as \\n and \\t in the
 * input xml file may mess up the parsing.
 */
void XLXmlFile::setXmlData(const std::string& xmlData)    // NOLINT
{
    Expects(m_xmlData != nullptr);
    m_xmlData->setRawData(xmlData);
}

/**
 * @details This method retrieves the underlying XML data as a std::string.
 */
std::string XLXmlFile::xmlData(XLXmlSavingDeclaration savingDeclaration) const
{
    Expects(m_xmlData != nullptr);
    return m_xmlData->getRawData(savingDeclaration);
}

/**
 * @details
 */
const XLDocument& XLXmlFile::parentDoc() const
{
    Expects(m_xmlData != nullptr);
    return *m_xmlData->getParentDoc();
}

/**
 * @details
 */
XLDocument& XLXmlFile::parentDoc()
{
    Expects(m_xmlData != nullptr);
    return *m_xmlData->getParentDoc();
}

/**
 * @details
 */
std::string XLXmlFile::relationshipID() const
{
    Expects(m_xmlData != nullptr);
    return m_xmlData->getXmlID();
}

/**
 * @details This method returns a pointer to the underlying XMLDocument resource.
 */
XMLDocument& XLXmlFile::xmlDocument()
{
    Expects(m_xmlData != nullptr);
    return const_cast<XMLDocument&>(static_cast<const XLXmlFile*>(this)->xmlDocument());    // NOLINT
    // return *m_xmlData->getXmlDocument();    // <- why not this easy version?
}

/**
 * @details This method returns a pointer to the underlying XMLDocument resource as const.
 */
const XMLDocument& XLXmlFile::xmlDocument() const
{
    Expects(m_xmlData != nullptr);
    return *m_xmlData->getXmlDocument();
}

/**
 * @details provide access to the underlying XLXmlData::getXmlPath() function
 */
std::string XLXmlFile::getXmlPath() const { return m_xmlData == nullptr ? "" : m_xmlData->getXmlPath(); }
