#include "XLConditionalFormatting.hpp"
#include "XLException.hpp"
#include "XLUtilities.hpp"
#include "XLWorksheet.hpp"
#include <algorithm>
#include <fmt/format.h>
#include <limits>
#include <map>

using namespace OpenXLSX;

namespace OpenXLSX
{
    /**
     * @brief get the correct XLCfType from the OOXML cfRule type attribute string
     * @param typeString the string as used in the OOXML
     * @return the corresponding XLCfType enum value
     */
    XLCfType XLCfTypeFromString(std::string const& typeString)
    {
        if (typeString == "expression") return XLCfType::Expression;
        if (typeString == "cellIs") return XLCfType::CellIs;
        if (typeString == "colorScale") return XLCfType::ColorScale;
        if (typeString == "dataBar") return XLCfType::DataBar;
        if (typeString == "iconSet") return XLCfType::IconSet;
        if (typeString == "top10") return XLCfType::Top10;
        if (typeString == "uniqueValues") return XLCfType::UniqueValues;
        if (typeString == "duplicateValues") return XLCfType::DuplicateValues;
        if (typeString == "containsText") return XLCfType::ContainsText;
        if (typeString == "notContainsText") return XLCfType::NotContainsText;
        if (typeString == "beginsWith") return XLCfType::BeginsWith;
        if (typeString == "endsWith") return XLCfType::EndsWith;
        if (typeString == "containsBlanks") return XLCfType::ContainsBlanks;
        if (typeString == "notContainsBlanks") return XLCfType::NotContainsBlanks;
        if (typeString == "containsErrors") return XLCfType::ContainsErrors;
        if (typeString == "notContainsErrors") return XLCfType::NotContainsErrors;
        if (typeString == "timePeriod") return XLCfType::TimePeriod;
        if (typeString == "aboveAverage") return XLCfType::AboveAverage;
        return XLCfType::Invalid;
    }

    /**
     * @brief inverse of XLCfTypeFromString
     * @param cfType the type for which to get the OOXML string
     */
    std::string XLCfTypeToString(XLCfType cfType)
    {
        switch (cfType) {
            case XLCfType::Expression:
                return "expression";
            case XLCfType::CellIs:
                return "cellIs";
            case XLCfType::ColorScale:
                return "colorScale";
            case XLCfType::DataBar:
                return "dataBar";
            case XLCfType::IconSet:
                return "iconSet";
            case XLCfType::Top10:
                return "top10";
            case XLCfType::UniqueValues:
                return "uniqueValues";
            case XLCfType::DuplicateValues:
                return "duplicateValues";
            case XLCfType::ContainsText:
                return "containsText";
            case XLCfType::NotContainsText:
                return "notContainsText";
            case XLCfType::BeginsWith:
                return "beginsWith";
            case XLCfType::EndsWith:
                return "endsWith";
            case XLCfType::ContainsBlanks:
                return "containsBlanks";
            case XLCfType::NotContainsBlanks:
                return "notContainsBlanks";
            case XLCfType::ContainsErrors:
                return "containsErrors";
            case XLCfType::NotContainsErrors:
                return "notContainsErrors";
            case XLCfType::TimePeriod:
                return "timePeriod";
            case XLCfType::AboveAverage:
                return "aboveAverage";
            case XLCfType::Invalid:
                [[fallthrough]];
            default:
                return "(invalid)";
        }
    }

    /**
     * @brief get the correct XLCfOperator from the OOXML cfRule operator attribute string
     * @param operatorString the string as used in the OOXML
     * @return the corresponding XLCfOperator enum value
     */
    XLCfOperator XLCfOperatorFromString(std::string const& operatorString)
    {
        if (operatorString == "lessThan") return XLCfOperator::LessThan;
        if (operatorString == "lessThanOrEqual") return XLCfOperator::LessThanOrEqual;
        if (operatorString == "equal") return XLCfOperator::Equal;
        if (operatorString == "notEqual") return XLCfOperator::NotEqual;
        if (operatorString == "greaterThanOrEqual") return XLCfOperator::GreaterThanOrEqual;
        if (operatorString == "greaterThan") return XLCfOperator::GreaterThan;
        if (operatorString == "between") return XLCfOperator::Between;
        if (operatorString == "notBetween") return XLCfOperator::NotBetween;
        if (operatorString == "containsText") return XLCfOperator::ContainsText;
        if (operatorString == "notContains") return XLCfOperator::NotContains;
        if (operatorString == "beginsWith") return XLCfOperator::BeginsWith;
        if (operatorString == "endsWith") return XLCfOperator::EndsWith;
        return XLCfOperator::Invalid;
    }

    /**
     * @brief inverse of XLCfOperatorFromString
     * @param cfOperator the XLCfOperator for which to get the OOXML string
     */
    std::string XLCfOperatorToString(XLCfOperator cfOperator)
    {
        switch (cfOperator) {
            case XLCfOperator::LessThan:
                return "lessThan";
            case XLCfOperator::LessThanOrEqual:
                return "lessThanOrEqual";
            case XLCfOperator::Equal:
                return "equal";
            case XLCfOperator::NotEqual:
                return "notEqual";
            case XLCfOperator::GreaterThanOrEqual:
                return "greaterThanOrEqual";
            case XLCfOperator::GreaterThan:
                return "greaterThan";
            case XLCfOperator::Between:
                return "between";
            case XLCfOperator::NotBetween:
                return "notBetween";
            case XLCfOperator::ContainsText:
                return "containsText";
            case XLCfOperator::NotContains:
                return "notContains";
            case XLCfOperator::BeginsWith:
                return "beginsWith";
            case XLCfOperator::EndsWith:
                return "endsWith";
            case XLCfOperator::Invalid:
                [[fallthrough]];
            default:
                return "(invalid)";
        }
    }

    /**
     * @brief get the correct XLCfTimePeriod from the OOXML cfRule timePeriod attribute string
     * @param timePeriodString the string as used in the OOXML
     * @return the corresponding XLCfTimePeriod enum value
     */
    XLCfTimePeriod XLCfTimePeriodFromString(std::string const& timePeriodString)
    {
        if (timePeriodString == "today") return XLCfTimePeriod::Today;
        if (timePeriodString == "yesterday") return XLCfTimePeriod::Yesterday;
        if (timePeriodString == "tomorrow") return XLCfTimePeriod::Tomorrow;
        if (timePeriodString == "last7Days") return XLCfTimePeriod::Last7Days;
        if (timePeriodString == "thisMonth") return XLCfTimePeriod::ThisMonth;
        if (timePeriodString == "lastMonth") return XLCfTimePeriod::LastMonth;
        if (timePeriodString == "nextMonth") return XLCfTimePeriod::NextMonth;
        if (timePeriodString == "thisWeek") return XLCfTimePeriod::ThisWeek;
        if (timePeriodString == "lastWeek") return XLCfTimePeriod::LastWeek;
        if (timePeriodString == "nextWeek") return XLCfTimePeriod::NextWeek;
        return XLCfTimePeriod::Invalid;
    }

    /**
     * @brief inverse of XLCfTimePeriodFromString
     * @param cfTimePeriod the XLCfTimePeriod for which to get the OOXML string
     */
    std::string XLCfTimePeriodToString(XLCfTimePeriod cfTimePeriod)
    {
        switch (cfTimePeriod) {
            case XLCfTimePeriod::Today:
                return "today";
            case XLCfTimePeriod::Yesterday:
                return "yesterday";
            case XLCfTimePeriod::Tomorrow:
                return "tomorrow";
            case XLCfTimePeriod::Last7Days:
                return "last7Days";
            case XLCfTimePeriod::ThisMonth:
                return "thisMonth";
            case XLCfTimePeriod::LastMonth:
                return "lastMonth";
            case XLCfTimePeriod::NextMonth:
                return "nextMonth";
            case XLCfTimePeriod::ThisWeek:
                return "thisWeek";
            case XLCfTimePeriod::LastWeek:
                return "lastWeek";
            case XLCfTimePeriod::NextWeek:
                return "nextWeek";
            case XLCfTimePeriod::Invalid:
                [[fallthrough]];
            default:
                return "(invalid)";
        }
    }

    XLCfvoType XLCfvoTypeFromString(std::string const& cfvoTypeString)
    {
        if (cfvoTypeString == "min") return XLCfvoType::Min;
        if (cfvoTypeString == "max") return XLCfvoType::Max;
        if (cfvoTypeString == "num") return XLCfvoType::Number;
        if (cfvoTypeString == "percent") return XLCfvoType::Percent;
        if (cfvoTypeString == "formula") return XLCfvoType::Formula;
        if (cfvoTypeString == "percentile") return XLCfvoType::Percentile;
        return XLCfvoType::Invalid;
    }

    std::string XLCfvoTypeToString(XLCfvoType cfvoType)
    {
        switch (cfvoType) {
            case XLCfvoType::Min:
                return "min";
            case XLCfvoType::Max:
                return "max";
            case XLCfvoType::Number:
                return "num";
            case XLCfvoType::Percent:
                return "percent";
            case XLCfvoType::Formula:
                return "formula";
            case XLCfvoType::Percentile:
                return "percentile";
            default:
                return "num";
        }
    }
}    // namespace OpenXLSX

XLCfvo::XLCfvo() : m_xmlDocument(std::make_unique<XMLDocument>()) { m_cfvoNode = m_xmlDocument->append_child("cfvo"); }
XLCfvo::XLCfvo(const XMLNode& node) : m_cfvoNode(node) {}
XLCfvo::XLCfvo(const XLCfvo& other) : m_cfvoNode(other.m_cfvoNode)
{
    if (other.m_xmlDocument) {
        m_xmlDocument = std::make_unique<XMLDocument>();
        m_xmlDocument->reset(*other.m_xmlDocument);
        m_cfvoNode = m_xmlDocument->document_element();
    }
}
XLCfvo::XLCfvo(XLCfvo&& other) noexcept : m_xmlDocument(std::move(other.m_xmlDocument)), m_cfvoNode(other.m_cfvoNode) {}
XLCfvo::~XLCfvo() = default;
XLCfvo& XLCfvo::operator=(const XLCfvo& other)
{
    if (this != &other) {
        m_cfvoNode = other.m_cfvoNode;
        if (other.m_xmlDocument) {
            m_xmlDocument = std::make_unique<XMLDocument>();
            m_xmlDocument->reset(*other.m_xmlDocument);
            m_cfvoNode = m_xmlDocument->document_element();
        }
        else {
            m_xmlDocument.reset();
        }
    }
    return *this;
}
XLCfvo& XLCfvo::operator=(XLCfvo&& other) noexcept
{
    if (this != &other) {
        m_xmlDocument = std::move(other.m_xmlDocument);
        m_cfvoNode    = other.m_cfvoNode;
    }
    return *this;
}

XLCfvoType  XLCfvo::type() const { return XLCfvoTypeFromString(m_cfvoNode.attribute("type").value()); }
std::string XLCfvo::value() const { return m_cfvoNode.attribute("val").value(); }
bool        XLCfvo::gte() const { return m_cfvoNode.attribute("gte").as_bool(true); }
void        XLCfvo::setType(XLCfvoType type)
{
    if (m_cfvoNode.attribute("type").empty()) m_cfvoNode.append_attribute("type");
    m_cfvoNode.attribute("type").set_value(XLCfvoTypeToString(type).c_str());
}
void XLCfvo::setValue(const std::string& value)
{
    if (value.empty()) { m_cfvoNode.remove_attribute("val"); }
    else {
        if (m_cfvoNode.attribute("val").empty()) m_cfvoNode.append_attribute("val");
        m_cfvoNode.attribute("val").set_value(value.c_str());
    }
}
void XLCfvo::setGte(bool gte)
{
    if (m_cfvoNode.attribute("gte").empty()) m_cfvoNode.append_attribute("gte");
    m_cfvoNode.attribute("gte").set_value(gte ? "1" : "0");
}

XLCfColorScale::XLCfColorScale() : m_xmlDocument(std::make_unique<XMLDocument>())
{ m_colorScaleNode = m_xmlDocument->append_child("colorScale"); }
XLCfColorScale::XLCfColorScale(const XMLNode& node) : m_colorScaleNode(node) {}
XLCfColorScale::XLCfColorScale(const XLCfColorScale& other) : m_colorScaleNode(other.m_colorScaleNode)
{
    if (other.m_xmlDocument) {
        m_xmlDocument = std::make_unique<XMLDocument>();
        m_xmlDocument->reset(*other.m_xmlDocument);
        m_colorScaleNode = m_xmlDocument->document_element();
    }
}
XLCfColorScale::XLCfColorScale(XLCfColorScale&& other) noexcept
    : m_xmlDocument(std::move(other.m_xmlDocument)),
      m_colorScaleNode(other.m_colorScaleNode)
{}
XLCfColorScale::~XLCfColorScale() = default;
XLCfColorScale& XLCfColorScale::operator=(const XLCfColorScale& other)
{
    if (this != &other) {
        m_colorScaleNode = other.m_colorScaleNode;
        if (other.m_xmlDocument) {
            m_xmlDocument = std::make_unique<XMLDocument>();
            m_xmlDocument->reset(*other.m_xmlDocument);
            m_colorScaleNode = m_xmlDocument->document_element();
        }
        else {
            m_xmlDocument.reset();
        }
    }
    return *this;
}
XLCfColorScale& XLCfColorScale::operator=(XLCfColorScale&& other) noexcept
{
    if (this != &other) {
        m_xmlDocument    = std::move(other.m_xmlDocument);
        m_colorScaleNode = other.m_colorScaleNode;
    }
    return *this;
}

std::vector<XLCfvo> XLCfColorScale::cfvos() const
{
    std::vector<XLCfvo> result;
    for (auto& node : m_colorScaleNode.children("cfvo")) result.emplace_back(node);
    return result;
}
std::vector<XLColor> XLCfColorScale::colors() const
{
    std::vector<XLColor> result;
    for (auto& node : m_colorScaleNode.children("color")) {
        auto attr = node.attribute("rgb");
        if (!attr.empty() && std::string_view(attr.value()).size() >= 6)
            result.emplace_back(attr.value());
        else
            result.emplace_back();
    }
    return result;
}
void XLCfColorScale::addValue(XLCfvoType type, const std::string& value, const XLColor& color)
{
    auto cfvo = m_colorScaleNode.append_child("cfvo");
    cfvo.append_attribute("type").set_value(XLCfvoTypeToString(type).c_str());
    if (!value.empty()) cfvo.append_attribute("val").set_value(value.c_str());

    auto colorNode = m_colorScaleNode.append_child("color");
    colorNode.append_attribute("rgb").set_value(color.hex().c_str());
}
void XLCfColorScale::clear() { m_colorScaleNode.remove_children(); }

XLCfDataBar::XLCfDataBar() : m_xmlDocument(std::make_unique<XMLDocument>()) { m_dataBarNode = m_xmlDocument->append_child("dataBar"); }
XLCfDataBar::XLCfDataBar(const XMLNode& node) : m_dataBarNode(node) {}
XLCfDataBar::XLCfDataBar(const XLCfDataBar& other) : m_dataBarNode(other.m_dataBarNode)
{
    if (other.m_xmlDocument) {
        m_xmlDocument = std::make_unique<XMLDocument>();
        m_xmlDocument->reset(*other.m_xmlDocument);
        m_dataBarNode = m_xmlDocument->document_element();
    }
}
XLCfDataBar::XLCfDataBar(XLCfDataBar&& other) noexcept : m_xmlDocument(std::move(other.m_xmlDocument)), m_dataBarNode(other.m_dataBarNode)
{}
XLCfDataBar::~XLCfDataBar() = default;
XLCfDataBar& XLCfDataBar::operator=(const XLCfDataBar& other)
{
    if (this != &other) {
        m_dataBarNode = other.m_dataBarNode;
        if (other.m_xmlDocument) {
            m_xmlDocument = std::make_unique<XMLDocument>();
            m_xmlDocument->reset(*other.m_xmlDocument);
            m_dataBarNode = m_xmlDocument->document_element();
        }
        else {
            m_xmlDocument.reset();
        }
    }
    return *this;
}
XLCfDataBar& XLCfDataBar::operator=(XLCfDataBar&& other) noexcept
{
    if (this != &other) {
        m_xmlDocument = std::move(other.m_xmlDocument);
        m_dataBarNode = other.m_dataBarNode;
    }
    return *this;
}

XLCfvo  XLCfDataBar::min() const { return XLCfvo(m_dataBarNode.child("cfvo")); }
XLCfvo  XLCfDataBar::max() const { return XLCfvo(m_dataBarNode.child("cfvo").next_sibling("cfvo")); }
XLColor XLCfDataBar::color() const
{
    auto attr = m_dataBarNode.child("color").attribute("rgb");
    return (!attr.empty() && std::string_view(attr.value()).size() >= 6) ? XLColor(attr.value()) : XLColor();
}
void XLCfDataBar::setMin(XLCfvoType type, const std::string& value)
{
    auto node = m_dataBarNode.child("cfvo");
    if (node.empty()) node = m_dataBarNode.prepend_child("cfvo");
    XLCfvo(node).setType(type);
    XLCfvo(node).setValue(value);
}
void XLCfDataBar::setMax(XLCfvoType type, const std::string& value)
{
    auto node = m_dataBarNode.child("cfvo").next_sibling("cfvo");
    if (node.empty()) node = m_dataBarNode.insert_child_after("cfvo", m_dataBarNode.child("cfvo"));
    XLCfvo(node).setType(type);
    XLCfvo(node).setValue(value);
}
void XLCfDataBar::setColor(const XLColor& color)
{
    auto node = m_dataBarNode.child("color");
    if (node.empty()) node = m_dataBarNode.append_child("color");
    if (node.attribute("rgb").empty()) node.append_attribute("rgb");
    node.attribute("rgb").set_value(color.hex().c_str());
}
bool XLCfDataBar::showValue() const { return m_dataBarNode.attribute("showValue").as_bool(true); }
void XLCfDataBar::setShowValue(bool show)
{
    if (m_dataBarNode.attribute("showValue").empty()) m_dataBarNode.append_attribute("showValue");
    m_dataBarNode.attribute("showValue").set_value(show ? "1" : "0");
}

XLCfIconSet::XLCfIconSet() : m_xmlDocument(std::make_unique<XMLDocument>()) { m_iconSetNode = m_xmlDocument->append_child("iconSet"); }
XLCfIconSet::XLCfIconSet(const XMLNode& node) : m_iconSetNode(node) {}
XLCfIconSet::XLCfIconSet(const XLCfIconSet& other) : m_iconSetNode(other.m_iconSetNode)
{
    if (other.m_xmlDocument) {
        m_xmlDocument = std::make_unique<XMLDocument>();
        m_xmlDocument->reset(*other.m_xmlDocument);
        m_iconSetNode = m_xmlDocument->document_element();
    }
}
XLCfIconSet::XLCfIconSet(XLCfIconSet&& other) noexcept : m_xmlDocument(std::move(other.m_xmlDocument)), m_iconSetNode(other.m_iconSetNode)
{}
XLCfIconSet::~XLCfIconSet() = default;
XLCfIconSet& XLCfIconSet::operator=(const XLCfIconSet& other)
{
    if (this != &other) {
        m_iconSetNode = other.m_iconSetNode;
        if (other.m_xmlDocument) {
            m_xmlDocument = std::make_unique<XMLDocument>();
            m_xmlDocument->reset(*other.m_xmlDocument);
            m_iconSetNode = m_xmlDocument->document_element();
        }
        else {
            m_xmlDocument.reset();
        }
    }
    return *this;
}
XLCfIconSet& XLCfIconSet::operator=(XLCfIconSet&& other) noexcept
{
    if (this != &other) {
        m_xmlDocument = std::move(other.m_xmlDocument);
        m_iconSetNode = other.m_iconSetNode;
    }
    return *this;
}

std::string XLCfIconSet::iconSet() const { return m_iconSetNode.attribute("iconSet").value(); }
void        XLCfIconSet::setIconSet(const std::string& iconSetName)
{
    if (m_iconSetNode.attribute("iconSet").empty()) m_iconSetNode.append_attribute("iconSet");
    m_iconSetNode.attribute("iconSet").set_value(iconSetName.c_str());
}
std::vector<XLCfvo> XLCfIconSet::cfvos() const
{
    std::vector<XLCfvo> result;
    for (auto& node : m_iconSetNode.children("cfvo")) result.emplace_back(node);
    return result;
}
void XLCfIconSet::addValue(XLCfvoType type, const std::string& value)
{
    auto cfvo = m_iconSetNode.append_child("cfvo");
    cfvo.append_attribute("type").set_value(XLCfvoTypeToString(type).c_str());
    if (!value.empty()) cfvo.append_attribute("val").set_value(value.c_str());
}
void XLCfIconSet::clear() { m_iconSetNode.remove_children(); }
bool XLCfIconSet::showValue() const { return m_iconSetNode.attribute("showValue").as_bool(true); }
void XLCfIconSet::setShowValue(bool show)
{
    if (m_iconSetNode.attribute("showValue").empty()) m_iconSetNode.append_attribute("showValue");
    m_iconSetNode.attribute("showValue").set_value(show ? "1" : "0");
}
bool XLCfIconSet::percent() const { return m_iconSetNode.attribute("percent").as_bool(true); }
void XLCfIconSet::setPercent(bool percent)
{
    if (m_iconSetNode.attribute("percent").empty()) m_iconSetNode.append_attribute("percent");
    m_iconSetNode.attribute("percent").set_value(percent ? "1" : "0");
}
bool XLCfIconSet::reverse() const { return m_iconSetNode.attribute("reverse").as_bool(false); }
void XLCfIconSet::setReverse(bool reverse)
{
    if (m_iconSetNode.attribute("reverse").empty()) m_iconSetNode.append_attribute("reverse");
    m_iconSetNode.attribute("reverse").set_value(reverse ? "1" : "0");
}

XLCfRule::XLCfRule() : m_xmlDocument(std::make_unique<XMLDocument>()), m_cfRuleNode(XMLNode())
{ m_cfRuleNode = m_xmlDocument->append_child("cfRule"); }
XLCfRule::XLCfRule(const XMLNode& node) : m_cfRuleNode(node) {}
XLCfRule::XLCfRule(const XLCfRule& other) : m_cfRuleNode(other.m_cfRuleNode)
{
    if (other.m_xmlDocument) {
        m_xmlDocument = std::make_unique<XMLDocument>();
        m_xmlDocument->reset(*other.m_xmlDocument);
        m_cfRuleNode = m_xmlDocument->document_element();
    }
}
XLCfRule::XLCfRule(XLCfRule&& other) noexcept : m_xmlDocument(std::move(other.m_xmlDocument)), m_cfRuleNode(other.m_cfRuleNode) {}
XLCfRule::~XLCfRule() = default;

XLCfRule& XLCfRule::operator=(const XLCfRule& other)
{
    if (&other != this) {
        m_cfRuleNode = other.m_cfRuleNode;
        if (other.m_xmlDocument) {
            m_xmlDocument = std::make_unique<XMLDocument>();
            m_xmlDocument->reset(*other.m_xmlDocument);
            m_cfRuleNode = m_xmlDocument->document_element();
        }
        else {
            m_xmlDocument.reset();
        }
    }
    return *this;
}

XLCfRule& XLCfRule::operator=(XLCfRule&& other) noexcept
{
    if (this != &other) {
        m_xmlDocument = std::move(other.m_xmlDocument);
        m_cfRuleNode  = other.m_cfRuleNode;
    }
    return *this;
}

bool XLCfRule::empty() const { return m_cfRuleNode.empty(); }

std::string XLCfRule::formula() const { return m_cfRuleNode.child("formula").first_child_of_type(pugi::node_pcdata).value(); }

std::vector<std::string> XLCfRule::formulas() const
{
    std::vector<std::string> result;
    for (auto& node : m_cfRuleNode.children("formula")) { result.emplace_back(node.text().get()); }
    return result;
}

XLCfColorScale       XLCfRule::colorScale() const { return XLCfColorScale(m_cfRuleNode.child("colorScale")); }
XLCfDataBar          XLCfRule::dataBar() const { return XLCfDataBar(m_cfRuleNode.child("dataBar")); }
XLCfIconSet          XLCfRule::iconSet() const { return XLCfIconSet(m_cfRuleNode.child("iconSet")); }
XLUnsupportedElement XLCfRule::extLst() const { return XLUnsupportedElement{}; }

XLCfType       XLCfRule::type() const { return XLCfTypeFromString(m_cfRuleNode.attribute("type").value()); }
XLStyleIndex   XLCfRule::dxfId() const { return m_cfRuleNode.attribute("dxfId").as_uint(XLInvalidStyleIndex); }
uint16_t       XLCfRule::priority() const { return static_cast<uint16_t>(m_cfRuleNode.attribute("priority").as_uint(XLPriorityNotSet)); }
bool           XLCfRule::stopIfTrue() const { return m_cfRuleNode.attribute("stopIfTrue").as_bool(false); }
bool           XLCfRule::aboveAverage() const { return m_cfRuleNode.attribute("aboveAverage").as_bool(false); }
bool           XLCfRule::percent() const { return m_cfRuleNode.attribute("percent").as_bool(false); }
bool           XLCfRule::bottom() const { return m_cfRuleNode.attribute("bottom").as_bool(false); }
XLCfOperator   XLCfRule::Operator() const { return XLCfOperatorFromString(m_cfRuleNode.attribute("operator").value()); }
std::string    XLCfRule::text() const { return m_cfRuleNode.attribute("text").value(); }
XLCfTimePeriod XLCfRule::timePeriod() const { return XLCfTimePeriodFromString(m_cfRuleNode.attribute("timePeriod").value()); }
uint16_t       XLCfRule::rank() const { return static_cast<uint16_t>(m_cfRuleNode.attribute("rank").as_uint()); }
int16_t        XLCfRule::stdDev() const { return static_cast<int16_t>(m_cfRuleNode.attribute("stdDev").as_int()); }
bool           XLCfRule::equalAverage() const { return m_cfRuleNode.attribute("equalAverage").as_bool(false); }

XLCfRule& XLCfRule::setFormula(std::string const& newFormula)
{
    clearFormulas();
    addFormula(newFormula);
    return *this;
}

bool XLCfRule::addFormula(std::string const& newFormula)
{
    // Find where to insert: before the first non-formula child that is in m_nodeOrder
    XMLNode insertBefore = XMLNode();
    for (auto child : m_cfRuleNode.children()) {
        std::string_view name = child.name();
        if (name != "formula") {
            insertBefore = child;
            break;
        }
    }

    XMLNode newNode =
        insertBefore.empty() ? m_cfRuleNode.append_child("formula") : m_cfRuleNode.insert_child_before("formula", insertBefore);
    return newNode.append_child(pugi::node_pcdata).set_value(newFormula.c_str());
}

void XLCfRule::clearFormulas()
{
    while (!m_cfRuleNode.child("formula").empty()) { m_cfRuleNode.remove_child("formula"); }
}

bool XLCfRule::setColorScale(XLCfColorScale const& newColorScale)
{
    auto node = appendAndGetNode(m_cfRuleNode, "colorScale", m_nodeOrder);
    node.remove_children();
    for (auto& cfvo : newColorScale.cfvos()) node.append_copy(cfvo.node());
    for (auto& color : newColorScale.colors()) {
        auto colorNode = node.append_child("color");
        colorNode.append_attribute("rgb").set_value(color.hex().c_str());
    }
    return true;
}

bool XLCfRule::setDataBar(XLCfDataBar const& newDataBar)
{
    auto node = appendAndGetNode(m_cfRuleNode, "dataBar", m_nodeOrder);
    node.remove_children();
    node.append_copy(newDataBar.min().node());
    node.append_copy(newDataBar.max().node());
    auto colorNode = node.append_child("color");
    colorNode.append_attribute("rgb").set_value(newDataBar.color().hex().c_str());
    return true;
}

bool XLCfRule::setIconSet(XLCfIconSet const& newIconSet)
{
    auto node = appendAndGetNode(m_cfRuleNode, "iconSet", m_nodeOrder);
    node.remove_attributes();
    node.remove_children();

    // Copy all attributes
    for (pugi::xml_attribute attr = newIconSet.node().first_attribute(); attr; attr = attr.next_attribute()) {
        node.append_attribute(attr.name()).set_value(attr.value());
    }

    // Copy all children
    for (pugi::xml_node child = newIconSet.node().first_child(); child; child = child.next_sibling()) { node.append_copy(child); }
    return true;
}

bool XLCfRule::setExtLst(XLUnsupportedElement const& newExtLst)
{
    OpenXLSX::ignore(newExtLst);
    return false;
}

XLCfRule& XLCfRule::setType(XLCfType newType)
{
    appendAndSetAttribute(m_cfRuleNode, "type", XLCfTypeToString(newType)).empty();
    return *this;
}
XLCfRule& XLCfRule::setDxfId(XLStyleIndex newDxfId)
{
    appendAndSetAttribute(m_cfRuleNode, "dxfId", std::to_string(newDxfId)).empty();
    return *this;
}
XLCfRule& XLCfRule::setPriority(uint16_t newPriority)
{
    appendAndSetAttribute(m_cfRuleNode, "priority", std::to_string(newPriority)).empty();
    return *this;
}
XLCfRule& XLCfRule::setStopIfTrue(bool set)
{
    appendAndSetAttribute(m_cfRuleNode, "stopIfTrue", (set ? "1" : "0")).empty();
    return *this;
}
XLCfRule& XLCfRule::setAboveAverage(bool set)
{
    appendAndSetAttribute(m_cfRuleNode, "aboveAverage", (set ? "1" : "0")).empty();
    return *this;
}
XLCfRule& XLCfRule::setPercent(bool set)
{
    appendAndSetAttribute(m_cfRuleNode, "percent", (set ? "1" : "0")).empty();
    return *this;
}
XLCfRule& XLCfRule::setBottom(bool set)
{
    appendAndSetAttribute(m_cfRuleNode, "bottom", (set ? "1" : "0")).empty();
    return *this;
}
XLCfRule& XLCfRule::setOperator(XLCfOperator newOperator)
{
    appendAndSetAttribute(m_cfRuleNode, "operator", XLCfOperatorToString(newOperator)).empty();
    return *this;
}
XLCfRule& XLCfRule::setText(std::string const& newText)
{
    appendAndSetAttribute(m_cfRuleNode, "text", newText.c_str()).empty();
    return *this;
}
XLCfRule& XLCfRule::setTimePeriod(XLCfTimePeriod newTimePeriod)
{
    appendAndSetAttribute(m_cfRuleNode, "timePeriod", XLCfTimePeriodToString(newTimePeriod)).empty();
    return *this;
}
XLCfRule& XLCfRule::setRank(uint16_t newRank)
{
    appendAndSetAttribute(m_cfRuleNode, "rank", std::to_string(newRank)).empty();
    return *this;
}
XLCfRule& XLCfRule::setStdDev(int16_t newStdDev)
{
    appendAndSetAttribute(m_cfRuleNode, "stdDev", std::to_string(newStdDev)).empty();
    return *this;
}
XLCfRule& XLCfRule::setEqualAverage(bool set)
{
    appendAndSetAttribute(m_cfRuleNode, "equalAverage", (set ? "1" : "0")).empty();
    return *this;
}

std::string XLCfRule::summary() const
{
    XLStyleIndex dxfIndex = dxfId();
    return fmt::format("formula is {}, type is {}, dxfId is {}, priority is {}, stopIfTrue: {}, aboveAverage: {}, percent: {}, bottom: {}, "
                       "operator is {}, text is \"{}\", timePeriod is {}, rank is {}, stdDev is {}, equalAverage: {}",
                       formula(),
                       XLCfTypeToString(type()),
                       dxfIndex == XLInvalidStyleIndex ? "(invalid)" : std::to_string(dxfId()),
                       priority(),
                       stopIfTrue() ? "true" : "false",
                       aboveAverage() ? "true" : "false",
                       percent() ? "true" : "false",
                       bottom() ? "true" : "false",
                       XLCfOperatorToString(Operator()),
                       text(),
                       XLCfTimePeriodToString(timePeriod()),
                       rank(),
                       stdDev(),
                       equalAverage() ? "true" : "false");
}

XLCfRules::XLCfRules() : m_conditionalFormattingNode(XMLNode()) {}
XLCfRules::XLCfRules(const XMLNode& node) : m_conditionalFormattingNode(node) {}
XLCfRules::XLCfRules(const XLCfRules& other) : m_conditionalFormattingNode(other.m_conditionalFormattingNode) {}
XLCfRules::~XLCfRules() = default;

XLCfRules& XLCfRules::operator=(const XLCfRules& other)
{
    if (&other != this) m_conditionalFormattingNode = other.m_conditionalFormattingNode;
    return *this;
}

bool XLCfRules::empty() const { return m_conditionalFormattingNode.empty(); }

uint16_t XLCfRules::maxPriorityValue() const
{
    XMLNode node = m_conditionalFormattingNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule") node = node.next_sibling_of_type(pugi::node_element);
    uint16_t maxPriority = XLPriorityNotSet;
    while (not node.empty() and std::string(node.name()) == "cfRule") {
        maxPriority = std::max(maxPriority, XLCfRule(node).priority());
        node        = node.next_sibling_of_type(pugi::node_element);
    }
    return maxPriority;
}

bool XLCfRules::setPriority(size_t cfRuleIndex, uint16_t newPriority)
{
    XLCfRule affectedRule = cfRuleByIndex(cfRuleIndex);
    if (newPriority == affectedRule.priority()) return true;

    XMLNode node = m_conditionalFormattingNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule") node = node.next_sibling_of_type(pugi::node_element);

    XMLNode node2           = node;
    bool    newPriorityFree = true;
    while (newPriorityFree and not node2.empty() and std::string(node2.name()) == "cfRule") {
        if (XLCfRule(node2).priority() == newPriority) newPriorityFree = false;
        node2 = node2.next_sibling_of_type(pugi::node_element);
    }

    if (!newPriorityFree) {
        size_t index = 0;
        while (not node.empty() and std::string(node.name()) == "cfRule") {
            if (index != cfRuleIndex) {
                XLCfRule rule(node);
                uint16_t prio = rule.priority();
                if (prio >= newPriority) rule.setPriority(prio + 1);
            }
            node = node.next_sibling_of_type(pugi::node_element);
            ++index;
        }
    }
    affectedRule.setPriority(newPriority);
    return true;
}

void XLCfRules::renumberPriorities(uint16_t increment)
{
    if (increment == 0) throw XLException("XLCfRules::renumberPriorities: increment must not be 0");
    std::multimap<uint16_t, XLCfRule> rules;
    XMLNode                           node = m_conditionalFormattingNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule") node = node.next_sibling_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) == "cfRule") {
        XLCfRule rule(node);
        rules.insert(std::pair(rule.priority(), std::move(rule)));
        node = node.next_sibling_of_type(pugi::node_element);
    }
    if (rules.size() * increment > std::numeric_limits<uint16_t>::max()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCfRules::renumberPriorities: amount of rules "s + std::to_string(rules.size()) + " with given increment "s +
                          std::to_string(increment) + " exceeds max range of uint16_t"s);
    }
    uint16_t prio = 0;
    for (auto& [key, rule] : rules) {
        prio += increment;
        rule.setPriority(prio);
    }
    rules.clear();
}

size_t XLCfRules::count() const
{
    XMLNode node = m_conditionalFormattingNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule") node = node.next_sibling_of_type(pugi::node_element);
    size_t count = 0;
    while (not node.empty() and std::string(node.name()) == "cfRule") {
        ++count;
        node = node.next_sibling_of_type(pugi::node_element);
    }
    return count;
}

XLCfRule XLCfRules::cfRuleByIndex(size_t index) const
{
    XMLNode node = m_conditionalFormattingNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule") node = node.next_sibling_of_type(pugi::node_element);
    if (not node.empty()) {
        size_t count = 0;
        while (not node.empty() and std::string(node.name()) == "cfRule" and count < index) {
            ++count;
            node = node.next_sibling_of_type(pugi::node_element);
        }
        if (count == index and std::string(node.name()) == "cfRule") return XLCfRule(node);
    }
    using namespace std::literals::string_literals;
    throw XLException("XLCfRules::"s + __func__ + ": cfRule with index "s + std::to_string(index) + " does not exist");
}

size_t XLCfRules::create(XLCfRule copyFrom, std::string cfRulePrefix)
{
    uint16_t explicitPrio = XLPriorityNotSet;
    if (!copyFrom.empty() && copyFrom.priority() != XLPriorityNotSet) { explicitPrio = copyFrom.priority(); }

    uint16_t maxPrio = maxPriorityValue();
    if (explicitPrio == XLPriorityNotSet && maxPrio == std::numeric_limits<uint16_t>::max()) {
        using namespace std::literals::string_literals;
        throw XLException(
            "XLCfRules::"s + __func__ +
            ": can not create a new cfRule entry: no available priority value - please renumberPriorities or otherwise free up the highest value"s);
    }

    size_t  index = count();
    XMLNode newNode{};
    if (index == 0)
        newNode = appendAndGetNode(m_conditionalFormattingNode, "cfRule", m_nodeOrder);
    else {
        XMLNode lastCfRule = cfRuleByIndex(index - 1).m_cfRuleNode;
        if (not lastCfRule.empty()) newNode = m_conditionalFormattingNode.insert_child_after("cfRule", lastCfRule);
    }
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCfRules::"s + __func__ + ": failed to create a new cfRule entry");
    }

    if (!copyFrom.empty()) {
        for (pugi::xml_attribute attr = copyFrom.node().first_attribute(); attr; attr = attr.next_attribute()) {
            newNode.append_attribute(attr.name()).set_value(attr.value());
        }
        for (pugi::xml_node child = copyFrom.node().first_child(); child; child = child.next_sibling()) { newNode.append_copy(child); }
    }

    m_conditionalFormattingNode.insert_child_before(pugi::node_pcdata, newNode).set_value(cfRulePrefix.c_str());

    if (explicitPrio != XLPriorityNotSet) { cfRuleByIndex(index).setPriority(explicitPrio); }
    else {
        cfRuleByIndex(index).setPriority(maxPrio + 1);
    }

    return index;
}

std::string XLCfRules::summary() const
{
    size_t rulesCount = count();
    if (rulesCount == 0) return "(no cfRule entries)";
    std::string result = "";
    for (size_t idx = 0; idx < rulesCount; ++idx) {
        result += fmt::format("cfRule[{}] {}", idx, cfRuleByIndex(idx).summary());
        if (idx + 1 < rulesCount) result += ", ";
    }
    return result;
}

XLConditionalFormat::XLConditionalFormat() : m_conditionalFormattingNode(XMLNode()) {}
XLConditionalFormat::XLConditionalFormat(const XMLNode& node) : m_conditionalFormattingNode(node) {}
XLConditionalFormat::XLConditionalFormat(const XLConditionalFormat& other) : m_conditionalFormattingNode(other.m_conditionalFormattingNode)
{}
XLConditionalFormat::~XLConditionalFormat() = default;

XLConditionalFormat& XLConditionalFormat::operator=(const XLConditionalFormat& other)
{
    if (&other != this) m_conditionalFormattingNode = other.m_conditionalFormattingNode;
    return *this;
}

bool        XLConditionalFormat::empty() const { return m_conditionalFormattingNode.empty(); }
std::string XLConditionalFormat::sqref() const { return m_conditionalFormattingNode.attribute("sqref").value(); }
XLCfRules   XLConditionalFormat::cfRules() const { return XLCfRules(m_conditionalFormattingNode); }
bool        XLConditionalFormat::setSqref(std::string newSqref)
{ return appendAndSetAttribute(m_conditionalFormattingNode, "sqref", newSqref).empty() == false; }
bool XLConditionalFormat::setExtLst(XLUnsupportedElement const& newExtLst)
{
    OpenXLSX::ignore(newExtLst);
    return false;
}

std::string XLConditionalFormat::summary() const { return fmt::format("sqref is {}, cfRules: {}", sqref(), cfRules().summary()); }

XLConditionalFormats::XLConditionalFormats() : m_sheetNode(XMLNode()), m_nodeOrder(&XLWorksheetNodeOrder) {}
XLConditionalFormats::XLConditionalFormats(const XMLNode& sheet) : m_sheetNode(sheet), m_nodeOrder(&XLWorksheetNodeOrder) {}
XLConditionalFormats::~XLConditionalFormats() {}
XLConditionalFormats::XLConditionalFormats(const XLConditionalFormats& other)
    : m_sheetNode(other.m_sheetNode),
      m_nodeOrder(other.m_nodeOrder)
{}
XLConditionalFormats::XLConditionalFormats(XLConditionalFormats&& other) noexcept
    : m_sheetNode(std::move(other.m_sheetNode)),
      m_nodeOrder(other.m_nodeOrder)
{}

XLConditionalFormats& XLConditionalFormats::operator=(const XLConditionalFormats& other)
{
    if (&other != this) {
        m_sheetNode = other.m_sheetNode;
        m_nodeOrder = other.m_nodeOrder;
    }
    return *this;
}

bool XLConditionalFormats::empty() const { return m_sheetNode.empty(); }

size_t XLConditionalFormats::count() const
{
    XMLNode node = m_sheetNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "conditionalFormatting") node = node.next_sibling_of_type(pugi::node_element);
    size_t count = 0;
    while (not node.empty() and std::string(node.name()) == "conditionalFormatting") {
        ++count;
        node = node.next_sibling_of_type(pugi::node_element);
    }
    return count;
}

XLConditionalFormat XLConditionalFormats::conditionalFormatByIndex(size_t index) const
{
    XMLNode node = m_sheetNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "conditionalFormatting") node = node.next_sibling_of_type(pugi::node_element);
    if (not node.empty()) {
        size_t count = 0;
        while (not node.empty() and std::string(node.name()) == "conditionalFormatting" and count < index) {
            ++count;
            node = node.next_sibling_of_type(pugi::node_element);
        }
        if (count == index and std::string(node.name()) == "conditionalFormatting") return XLConditionalFormat(node);
    }
    using namespace std::literals::string_literals;
    throw XLException("XLConditionalFormats::"s + __func__ + ": conditional format with index "s + std::to_string(index) +
                      " does not exist");
}

size_t XLConditionalFormats::create(XLConditionalFormat copyFrom, std::string conditionalFormattingPrefix)
{
    size_t  index = count();
    XMLNode newNode{};
    if (index == 0)
        newNode = appendAndGetNode(m_sheetNode, "conditionalFormatting", *m_nodeOrder);
    else {
        XMLNode lastConditionalFormat = conditionalFormatByIndex(index - 1).m_conditionalFormattingNode;
        if (not lastConditionalFormat.empty()) newNode = m_sheetNode.insert_child_after("conditionalFormatting", lastConditionalFormat);
    }
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLConditionalFormats::"s + __func__ + ": failed to create a new conditional formatting entry");
    }

    if (!copyFrom.empty()) {
        for (pugi::xml_attribute attr = copyFrom.m_conditionalFormattingNode.first_attribute(); attr; attr = attr.next_attribute()) {
            newNode.append_attribute(attr.name()).set_value(attr.value());
        }
        for (pugi::xml_node child = copyFrom.m_conditionalFormattingNode.first_child(); child; child = child.next_sibling()) {
            newNode.append_copy(child);
        }
    }

    m_sheetNode.insert_child_before(pugi::node_pcdata, newNode).set_value(conditionalFormattingPrefix.c_str());
    return index;
}

std::string XLConditionalFormats::summary() const
{
    size_t conditionalFormatsCount = count();
    if (conditionalFormatsCount == 0) return "(no conditionalFormatting entries)";
    std::string result = "";
    for (size_t idx = 0; idx < conditionalFormatsCount; ++idx) {
        result += fmt::format("conditionalFormatting[{}] {}", idx, conditionalFormatByIndex(idx).summary());
        if (idx + 1 < conditionalFormatsCount) result += ", ";
    }
    return result;
}

// ----- Helper Builder Functions for XLCfRule -----

namespace OpenXLSX
{

    XLCfRule XLColorScaleRule(const XLColor& minColor, const XLColor& maxColor)
    {
        XLCfRule rule;
        rule.setType(XLCfType::ColorScale);
        XLCfColorScale colorScale;
        colorScale.addValue(XLCfvoType::Min, "", minColor);
        colorScale.addValue(XLCfvoType::Max, "", maxColor);
        rule.setColorScale(colorScale);
        return rule;
    }

    XLCfRule XLColorScaleRule(const XLColor& minColor, const XLColor& midColor, const XLColor& maxColor)
    {
        XLCfRule rule;
        rule.setType(XLCfType::ColorScale);
        XLCfColorScale colorScale;
        colorScale.addValue(XLCfvoType::Min, "", minColor);
        colorScale.addValue(XLCfvoType::Percentile, "50", midColor);
        colorScale.addValue(XLCfvoType::Max, "", maxColor);
        rule.setColorScale(colorScale);
        return rule;
    }

    XLCfRule XLDataBarRule(const XLColor& color, bool showValue)
    {
        XLCfRule rule;
        rule.setType(XLCfType::DataBar);
        XLCfDataBar dataBar;
        dataBar.setMin(XLCfvoType::Min, "");
        dataBar.setMax(XLCfvoType::Max, "");
        dataBar.setColor(color);
        dataBar.setShowValue(showValue);
        rule.setDataBar(dataBar);
        return rule;
    }

    XLCfRule XLCellIsRule(XLCfOperator op, const std::string& value)
    {
        XLCfRule rule;
        rule.setType(XLCfType::CellIs);
        rule.setOperator(op);
        rule.addFormula(value);
        return rule;
    }

    XLCfRule XLCellIsRule(const std::string& op, const std::string& value)
    {
        XLCfOperator cfOp = XLCfOperator::Equal;
        if (op == ">")
            cfOp = XLCfOperator::GreaterThan;
        else if (op == ">=")
            cfOp = XLCfOperator::GreaterThanOrEqual;
        else if (op == "<")
            cfOp = XLCfOperator::LessThan;
        else if (op == "<=")
            cfOp = XLCfOperator::LessThanOrEqual;
        else if (op == "==" || op == "=")
            cfOp = XLCfOperator::Equal;
        else if (op == "!=" || op == "<>")
            cfOp = XLCfOperator::NotEqual;
        else
            cfOp = XLCfOperatorFromString(op);

        return XLCellIsRule(cfOp, value);
    }

    XLCfRule XLFormulaRule(const std::string& formula)
    {
        XLCfRule rule;
        rule.setType(XLCfType::Expression);
        rule.addFormula(formula);
        return rule;
    }

    // Advanced builders
    XLCfRule XLIconSetRule(const std::string& iconSetName, bool showValue, bool reverse)
    {
        XLCfRule rule;
        rule.setType(XLCfType::IconSet);
        XLCfIconSet iconSet;
        iconSet.setIconSet(iconSetName);

        if (!showValue) iconSet.setShowValue(false);
        if (reverse) iconSet.setReverse(true);

        // Determine number of icons by looking at the first character of the iconSetName (e.g., "3Arrows", "4TrafficLights", "5Rating")
        int numIcons = 3;    // default
        if (!iconSetName.empty() && std::isdigit(iconSetName[0])) { numIcons = iconSetName[0] - '0'; }

        // Generate evenly spaced percentile thresholds (e.g., for 3: 0, 33, 67)
        for (int i = 0; i < numIcons; ++i) {
            int percent = (i * 100) / numIcons;
            iconSet.addValue(XLCfvoType::Percent, std::to_string(percent));
        }

        rule.setIconSet(iconSet);
        return rule;
    }

    XLCfRule XLTop10Rule(uint16_t rank, bool percent, bool bottom)
    {
        XLCfRule rule;
        rule.setType(XLCfType::Top10);
        rule.setRank(rank);
        if (percent) rule.setPercent(true);
        if (bottom) rule.setBottom(true);
        return rule;
    }

    XLCfRule XLAboveAverageRule(bool aboveAverage, bool equalAverage, int16_t stdDev)
    {
        XLCfRule rule;
        rule.setType(XLCfType::AboveAverage);
        if (!aboveAverage) rule.setAboveAverage(false);
        if (equalAverage) rule.setEqualAverage(true);
        if (stdDev > 0) rule.setStdDev(stdDev);
        return rule;
    }

    XLCfRule XLDuplicateValuesRule(bool unique)
    {
        XLCfRule rule;
        rule.setType(unique ? XLCfType::UniqueValues : XLCfType::DuplicateValues);
        return rule;
    }

    XLCfRule XLContainsTextRule(const std::string& text)
    {
        XLCfRule rule;
        rule.setType(XLCfType::ContainsText);
        rule.setText(text);
        rule.setOperator(XLCfOperator::ContainsText);
        return rule;
    }

    XLCfRule XLNotContainsTextRule(const std::string& text)
    {
        XLCfRule rule;
        rule.setType(XLCfType::NotContainsText);
        rule.setText(text);
        rule.setOperator(XLCfOperator::NotContains);
        return rule;
    }

    XLCfRule XLContainsBlanksRule()
    {
        XLCfRule rule;
        rule.setType(XLCfType::ContainsBlanks);
        return rule;
    }

    XLCfRule XLNotContainsBlanksRule()
    {
        XLCfRule rule;
        rule.setType(XLCfType::NotContainsBlanks);
        return rule;
    }

    XLCfRule XLContainsErrorsRule()
    {
        XLCfRule rule;
        rule.setType(XLCfType::ContainsErrors);
        return rule;
    }

    XLCfRule XLNotContainsErrorsRule()
    {
        XLCfRule rule;
        rule.setType(XLCfType::NotContainsErrors);
        return rule;
    }

}    // namespace OpenXLSX
