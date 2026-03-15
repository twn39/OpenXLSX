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

#ifndef OPENXLSX_XLPAGESETUP_HPP
#define OPENXLSX_XLPAGESETUP_HPP

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    /**
     * @brief Enum defining page orientation.
     */
    enum class XLPageOrientation { Default, Portrait, Landscape };

    /**
     * @brief A class representing the page margins of a worksheet.
     */
    class OPENXLSX_EXPORT XLPageMargins
    {
    public:
        XLPageMargins() = default;
        explicit XLPageMargins(const XMLNode& node);

        double left() const;
        void   setLeft(double value);

        double right() const;
        void   setRight(double value);

        double top() const;
        void   setTop(double value);

        double bottom() const;
        void   setBottom(double value);

        double header() const;
        void   setHeader(double value);

        double footer() const;
        void   setFooter(double value);

    private:
        XMLNode m_node;
    };

    /**
     * @brief A class representing the print options of a worksheet.
     */
    class OPENXLSX_EXPORT XLPrintOptions
    {
    public:
        XLPrintOptions() = default;
        explicit XLPrintOptions(const XMLNode& node);

        bool gridLines() const;
        void setGridLines(bool value);

        bool headings() const;
        void setHeadings(bool value);

        bool horizontalCentered() const;
        void setHorizontalCentered(bool value);

        bool verticalCentered() const;
        void setVerticalCentered(bool value);

    private:
        XMLNode m_node;
    };

    /**
     * @brief A class representing the page setup of a worksheet.
     */
    class OPENXLSX_EXPORT XLPageSetup
    {
    public:
        XLPageSetup() = default;
        explicit XLPageSetup(const XMLNode& node);

        uint32_t paperSize() const;
        void     setPaperSize(uint32_t value);

        XLPageOrientation orientation() const;
        void              setOrientation(XLPageOrientation value);

        uint32_t scale() const;
        void     setScale(uint32_t value);

        uint32_t fitToWidth() const;
        void     setFitToWidth(uint32_t value);

        uint32_t fitToHeight() const;
        void     setFitToHeight(uint32_t value);

        bool blackAndWhite() const;
        void setBlackAndWhite(bool value);

    private:
        XMLNode m_node;
    };

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLPAGESETUP_HPP
