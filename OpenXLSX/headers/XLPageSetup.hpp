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
