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
