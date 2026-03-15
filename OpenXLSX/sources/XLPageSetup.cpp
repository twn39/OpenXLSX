// ===== External Includes ===== //
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLPageSetup.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLPageMargins::XLPageMargins(const XMLNode& node) : m_node(node) {}

double XLPageMargins::left() const { return m_node.attribute("left").as_double(); }
void   XLPageMargins::setLeft(double value) { appendAndSetAttribute(m_node, "left", std::to_string(value)); }

double XLPageMargins::right() const { return m_node.attribute("right").as_double(); }
void   XLPageMargins::setRight(double value) { appendAndSetAttribute(m_node, "right", std::to_string(value)); }

double XLPageMargins::top() const { return m_node.attribute("top").as_double(); }
void   XLPageMargins::setTop(double value) { appendAndSetAttribute(m_node, "top", std::to_string(value)); }

double XLPageMargins::bottom() const { return m_node.attribute("bottom").as_double(); }
void   XLPageMargins::setBottom(double value) { appendAndSetAttribute(m_node, "bottom", std::to_string(value)); }

double XLPageMargins::header() const { return m_node.attribute("header").as_double(); }
void   XLPageMargins::setHeader(double value) { appendAndSetAttribute(m_node, "header", std::to_string(value)); }

double XLPageMargins::footer() const { return m_node.attribute("footer").as_double(); }
void   XLPageMargins::setFooter(double value) { appendAndSetAttribute(m_node, "footer", std::to_string(value)); }

/**
 * @details
 */
XLPrintOptions::XLPrintOptions(const XMLNode& node) : m_node(node) {}

bool XLPrintOptions::gridLines() const { return m_node.attribute("gridLines").as_bool(); }
void XLPrintOptions::setGridLines(bool value) { appendAndSetAttribute(m_node, "gridLines", value ? "1" : "0"); }

bool XLPrintOptions::headings() const { return m_node.attribute("headings").as_bool(); }
void XLPrintOptions::setHeadings(bool value) { appendAndSetAttribute(m_node, "headings", value ? "1" : "0"); }

bool XLPrintOptions::horizontalCentered() const { return m_node.attribute("horizontalCentered").as_bool(); }
void XLPrintOptions::setHorizontalCentered(bool value) { appendAndSetAttribute(m_node, "horizontalCentered", value ? "1" : "0"); }

bool XLPrintOptions::verticalCentered() const { return m_node.attribute("verticalCentered").as_bool(); }
void XLPrintOptions::setVerticalCentered(bool value) { appendAndSetAttribute(m_node, "verticalCentered", value ? "1" : "0"); }

/**
 * @details
 */
XLPageSetup::XLPageSetup(const XMLNode& node) : m_node(node) {}

uint32_t XLPageSetup::paperSize() const { return m_node.attribute("paperSize").as_uint(); }
void     XLPageSetup::setPaperSize(uint32_t value) { appendAndSetAttribute(m_node, "paperSize", std::to_string(value)); }

XLPageOrientation XLPageSetup::orientation() const
{
    std::string val = m_node.attribute("orientation").value();
    if (val == "portrait") return XLPageOrientation::Portrait;
    if (val == "landscape") return XLPageOrientation::Landscape;
    return XLPageOrientation::Default;
}

void XLPageSetup::setOrientation(XLPageOrientation value)
{
    if (value == XLPageOrientation::Portrait)
        appendAndSetAttribute(m_node, "orientation", "portrait");
    else if (value == XLPageOrientation::Landscape)
        appendAndSetAttribute(m_node, "orientation", "landscape");
    else
        m_node.remove_attribute("orientation");
}

uint32_t XLPageSetup::scale() const { return m_node.attribute("scale").as_uint(100); }
void     XLPageSetup::setScale(uint32_t value) { appendAndSetAttribute(m_node, "scale", std::to_string(value)); }

uint32_t XLPageSetup::fitToWidth() const { return m_node.attribute("fitToWidth").as_uint(1); }
void     XLPageSetup::setFitToWidth(uint32_t value) { appendAndSetAttribute(m_node, "fitToWidth", std::to_string(value)); }

uint32_t XLPageSetup::fitToHeight() const { return m_node.attribute("fitToHeight").as_uint(1); }
void     XLPageSetup::setFitToHeight(uint32_t value) { appendAndSetAttribute(m_node, "fitToHeight", std::to_string(value)); }

bool XLPageSetup::blackAndWhite() const { return m_node.attribute("blackAndWhite").as_bool(); }
void XLPageSetup::setBlackAndWhite(bool value) { appendAndSetAttribute(m_node, "blackAndWhite", value ? "1" : "0"); }
