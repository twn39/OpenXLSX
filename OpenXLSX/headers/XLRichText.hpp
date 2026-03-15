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

#ifndef OPENXLSX_XLRICHTEXT_HPP
#define OPENXLSX_XLRICHTEXT_HPP

// ===== External Includes ===== //
#include <optional>
#include <string>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLColor.hpp"

namespace OpenXLSX
{
    /**
     * @brief A class representing a single run of rich text.
     */
    class OPENXLSX_EXPORT XLRichTextRun
    {
    public:
        XLRichTextRun() = default;
        explicit XLRichTextRun(const std::string& text) : m_text(text) {}

        const std::string& text() const { return m_text; }
        void               setText(const std::string& text) { m_text = text; }

        // Font properties
        std::optional<std::string> fontName() const { return m_fontName; }
        void                       setFontName(const std::string& name) { m_fontName = name; }

        std::optional<size_t> fontSize() const { return m_fontSize; }
        void                  setFontSize(size_t size) { m_fontSize = size; }

        std::optional<XLColor> fontColor() const { return m_fontColor; }
        void                   setFontColor(const XLColor& color) { m_fontColor = color; }

        std::optional<bool> bold() const { return m_bold; }
        void                setBold(bool bold) { m_bold = bold; }

        std::optional<bool> italic() const { return m_italic; }
        void                setItalic(bool italic) { m_italic = italic; }

        std::optional<bool> underline() const { return m_underline; }
        void                setUnderline(bool underline) { m_underline = underline; }

        std::optional<bool> strikethrough() const { return m_strikethrough; }
        void                setStrikethrough(bool strikethrough) { m_strikethrough = strikethrough; }

    private:
        std::string                m_text;
        std::optional<std::string> m_fontName;
        std::optional<size_t>      m_fontSize;
        std::optional<XLColor>     m_fontColor;
        std::optional<bool>        m_bold;
        std::optional<bool>        m_italic;
        std::optional<bool>        m_underline;
        std::optional<bool>        m_strikethrough;
    };

    /**
     * @brief A class representing rich text in a cell.
     */
    class OPENXLSX_EXPORT XLRichText
    {
    public:
        XLRichText() = default;
        explicit XLRichText(const std::string& plainText) { m_runs.emplace_back(plainText); }

        void                              addRun(const XLRichTextRun& run) { m_runs.push_back(run); }
        const std::vector<XLRichTextRun>& runs() const { return m_runs; }
        std::vector<XLRichTextRun>&       runs() { return m_runs; }

        /**
         * @brief Get the plain text representation of the rich text.
         */
        std::string plainText() const
        {
            std::string result;
            for (const auto& run : m_runs) { result += run.text(); }
            return result;
        }

        bool empty() const { return m_runs.empty(); }
        void clear() { m_runs.clear(); }

    private:
        std::vector<XLRichTextRun> m_runs;
    };

    // Equality operators
    inline bool operator==(const XLRichTextRun& lhs, const XLRichTextRun& rhs)
    {
        return lhs.text() == rhs.text() && lhs.fontName() == rhs.fontName() && lhs.fontSize() == rhs.fontSize() &&
               lhs.fontColor() == rhs.fontColor() && lhs.bold() == rhs.bold() && lhs.italic() == rhs.italic() &&
               lhs.underline() == rhs.underline() && lhs.strikethrough() == rhs.strikethrough();
    }

    inline bool operator!=(const XLRichTextRun& lhs, const XLRichTextRun& rhs) { return !(lhs == rhs); }

    inline bool operator==(const XLRichText& lhs, const XLRichText& rhs) { return lhs.runs() == rhs.runs(); }

    inline bool operator!=(const XLRichText& lhs, const XLRichText& rhs) { return !(lhs == rhs); }

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLRICHTEXT_HPP
