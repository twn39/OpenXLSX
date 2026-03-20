#ifndef OPENXLSX_XLCHART_HPP
#define OPENXLSX_XLCHART_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include <string>
#include <string_view>

namespace OpenXLSX
{
    enum class XLChartType {
        Bar,
        Line,
        Pie,
        Scatter,
        Area,
        Doughnut
    };

    enum class XLLegendPosition {
        Bottom,
        Left,
        Right,
        Top,
        TopRight,
        Hidden
    };

    /**
     * @brief The XLChart class represents an Excel chart XML file.
     */
    class OPENXLSX_EXPORT XLChart final : public XLXmlFile
    {
    public:
        /**
         * @brief Constructor
         */
        XLChart() : XLXmlFile(nullptr) {}

        /**
         * @brief Constructor from XLXmlData
         * @param xmlData the source XML data
         */
        explicit XLChart(XLXmlData* xmlData);

        /**
         * @brief Copy Constructor
         * @param other The object to be copied
         */
        XLChart(const XLChart& other) = default;

        /**
         * @brief Move Constructor
         */
        XLChart(XLChart&& other) noexcept = default;

        /**
         * @brief Destructor
         */
        ~XLChart() = default;

        /**
         * @brief Copy Assignment Operator
         */
        XLChart& operator=(const XLChart& other) = default;

        /**
         * @brief Move Assignment Operator
         */
        XLChart& operator=(XLChart&& other) noexcept = default;

        /**
         * @brief Add a data series to the chart
         * @param valuesRef A cell reference to the data values (e.g. "Sheet1!$B$1:$B$10")
         * @param title A literal string or cell reference for the series name (e.g. "Revenue" or "Sheet1!$B$1")
         * @param categoriesRef A cell reference for the X-axis categories (e.g. "Sheet1!$A$1:$A$10")
         */
        void addSeries(std::string_view valuesRef, std::string_view title = "", std::string_view categoriesRef = "");

        /**
         * @brief Set the chart title
         * @param title The title text
         */
        void setTitle(std::string_view title);

        /**
         * @brief Set the legend position or hide it
         * @param position The legend position
         */
        void setLegendPosition(XLLegendPosition position);

    private:
        friend class XLDocument;
        void initXml(XLChartType type = XLChartType::Bar);
        [[nodiscard]] uint32_t seriesCount() const;
    };
}

#endif // OPENXLSX_XLCHART_HPP
