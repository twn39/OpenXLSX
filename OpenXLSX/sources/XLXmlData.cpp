// ===== External Includes ===== //
#include <pugixml.hpp>
#include <sstream>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLXmlData.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLXmlData::XLXmlData(XLDocument* parentDoc, const std::string& xmlPath, const std::string& xmlId, XLContentType xmlType)
    : m_parentDoc(parentDoc),
      m_xmlPath(xmlPath),
      m_xmlID(xmlId),
      m_xmlType(xmlType),
      m_xmlDoc(std::make_unique<XMLDocument>())
{ m_xmlDoc->reset(); }

/**
 * @details
 */
XLXmlData::~XLXmlData() = default;

/**
 * @details
 */
void XLXmlData::setRawData(const std::string& data)    // NOLINT
{ m_xmlDoc->load_string(data.c_str(), pugi_parse_settings); }

/**
 * @details
 * @note Default encoding for pugixml xml_document::save is pugi::encoding_auto, becomes pugi::encoding_utf8
 */
std::string XLXmlData::getRawData(XLXmlSavingDeclaration savingDeclaration) const
{
    XMLDocument* doc = const_cast<XMLDocument*>(getXmlDocument());

    // ===== 2024-07-08: ensure that the default encoding UTF-8 is explicitly written to the XML document with a custom saving declaration
    XMLNode saveDeclaration;
    for (auto node : doc->children()) {
        if (node.type() == pugi::node_declaration) {
            saveDeclaration = node;
            break;
        }
    }

    if (saveDeclaration.empty()) {    // if saving declaration node does not exist
        doc->prepend_child(pugi::node_pcdata).set_value("\n");                            // prepend a line break
        saveDeclaration = doc->prepend_child(pugi::node_declaration);                     // prepend a saving declaration
    }

    // ===== If a node_declaration could be fetched or created
    if (not saveDeclaration.empty()) {
        // ===== Fetch or create saving declaration attributes
        XMLAttribute attrVersion = saveDeclaration.attribute("version");
        if (attrVersion.empty()) attrVersion = saveDeclaration.append_attribute("version");
        XMLAttribute attrEncoding = saveDeclaration.attribute("encoding");
        if (attrEncoding.empty()) attrEncoding = saveDeclaration.append_attribute("encoding");
        XMLAttribute attrStandalone = saveDeclaration.attribute("standalone");
        if (savingDeclaration.standalone_as_bool()) {
            if (attrStandalone.empty()) attrStandalone = saveDeclaration.append_attribute("standalone");
            attrStandalone.set_value("yes");
        }
        else if (!attrStandalone.empty()) {
            saveDeclaration.remove_attribute(attrStandalone);
        }

        // ===== Set saving declaration attribute values (potentially overwriting existing values)
        attrVersion  = savingDeclaration.version().c_str();     // version="1.0" is XML default
        attrEncoding = savingDeclaration.encoding().c_str();    // encoding="UTF-8" is XML default
    }

    std::ostringstream ostr;
    doc->save(ostr, "", pugi::format_default);
    return ostr.str();
}

/**
 * @details
 */
XLDocument* XLXmlData::getParentDoc() { return m_parentDoc; }

/**
 * @details
 */
const XLDocument* XLXmlData::getParentDoc() const { return m_parentDoc; }

/**
 * @details
 */
std::string XLXmlData::getXmlPath() const { return m_xmlPath; }

/**
 * @details
 */
std::string XLXmlData::getXmlID() const { return m_xmlID; }

/**
 * @details
 */
XLContentType XLXmlData::getXmlType() const { return m_xmlType; }

/**
 * @details
 */
XMLDocument* XLXmlData::getXmlDocument()
{
    if (!m_xmlDoc->document_element()) m_xmlDoc->load_string(m_parentDoc->extractXmlFromArchive(m_xmlPath).c_str(), pugi_parse_settings);

    return m_xmlDoc.get();
}

/**
 * @details
 */
const XMLDocument* XLXmlData::getXmlDocument() const
{
    if (!m_xmlDoc->document_element()) m_xmlDoc->load_string(m_parentDoc->extractXmlFromArchive(m_xmlPath).c_str(), pugi_parse_settings);

    return m_xmlDoc.get();
}
