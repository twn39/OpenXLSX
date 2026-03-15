/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2026, Curry Tang

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

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
