#ifndef OPENXLSX_XLTABLES_HPP
#define OPENXLSX_XLTABLES_HPP

// ===== External Includes ===== //
#include <cstdint>    // uint8_t, uint16_t, uint32_t
#include <ostream>    // std::basic_ostream
// #include <type_traits>
// #include <variant>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
// #include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLXmlData.hpp"
#include "XLXmlFile.hpp"
#include "XLAutoFilter.hpp"
#include "XLTableColumn.hpp"

namespace OpenXLSX
{
    class XLWorksheet;

    /**
     * @brief The XLTables class is the base class for worksheet tables
     */
    class OPENXLSX_EXPORT XLTables : public XLXmlFile
    {
        friend class XLWorksheet;    // for access to XLXmlFile::getXmlPath
    public:
        /**
         * @brief Constructor
         */
        XLTables() : XLXmlFile(nullptr) {};

        /**
         * @brief The constructor.
         * @param xmlData the source XML of the table file
         */
        XLTables(XLXmlData* xmlData);

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         * @note The default copy constructor is used, i.e. only shallow copying of pointer data members.
         */
        XLTables(const XLTables& other) = default;

        /**
         * @brief
         * @param other
         */
        XLTables(XLTables&& other) noexcept = default;

        /**
         * @brief The destructor
         * @note The default destructor is used, since cleanup of pointer data members is not required.
         */
        ~XLTables() = default;

        /**
         * @brief Assignment operator
         * @return A reference to the new object.
         * @note The default assignment operator is used, i.e. only shallow copying of pointer data members.
         */
        XLTables& operator=(const XLTables&) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLTables& operator=(XLTables&& other) noexcept = default;

        /**
         * @brief Get the name of the table
         * @return std::string The table name
         */
        std::string name() const;

        /**
         * @brief Set the name of the table
         * @param name The new table name (must not contain spaces)
         */
        void setName(std::string_view name);

        /**
         * @brief Get the display name of the table
         * @return std::string The display name
         */
        std::string displayName() const;

        /**
         * @brief Set the display name of the table
         * @param name The new display name
         */
        void setDisplayName(std::string_view name);

        /**
         * @brief Get the range reference of the table (e.g., "A1:C10")
         * @return std::string The range reference
         */
        std::string rangeReference() const;

        /**
         * @brief Set the range reference of the table
         * @param ref The new range reference (e.g., "A1:C10")
         */
        void setRangeReference(std::string_view ref);

        /**
         * @brief Get the table style name
         * @return std::string The style name
         */
        std::string styleName() const;

        /**
         * @brief Set the table style name
         * @param styleName The new style name (e.g., "TableStyleMedium2")
         */
        void setStyleName(std::string_view styleName);

        /**
         * @brief Check if row stripes are shown
         * @return bool True if row stripes are shown
         */
        bool showRowStripes() const;

        /**
         * @brief Set whether row stripes are shown
         * @param show True to show row stripes
         */
        void setShowRowStripes(bool show);

        /**
         * @brief Check if the header row is shown
         * @return bool True if the header row is shown
         */
        bool showHeaderRow() const;

        /**
         * @brief Set whether the header row is shown
         * @param show True to show the header row
         */
        void setShowHeaderRow(bool show);

        /**
         * @brief Check if the totals row is shown
         * @return bool True if the totals row is shown
         */
        bool showTotalsRow() const;

        /**
         * @brief Set whether the totals row is shown
         * @param show True to show the totals row
         */
        void setShowTotalsRow(bool show);

        /**
         * @brief Check if column stripes are shown
         * @return bool True if column stripes are shown
         */
        bool showColumnStripes() const;

        /**
         * @brief Set whether column stripes are shown
         * @param show True to show column stripes
         */
        void setShowColumnStripes(bool show);

        /**
         * @brief Check if the first column is highlighted
         * @return bool True if the first column is highlighted
         */
        bool showFirstColumn() const;

        /**
         * @brief Set whether the first column is highlighted
         * @param show True to highlight the first column
         */
        void setShowFirstColumn(bool show);

        /**
         * @brief Check if the last column is highlighted
         * @return bool True if the last column is highlighted
         */
        bool showLastColumn() const;

        /**
         * @brief Set whether the last column is highlighted
         * @param show True to highlight the last column
         */
        void setShowLastColumn(bool show);

        /**
         * @brief Append a new column to the table
         * @param name The column name
         * @return The appended XLTableColumn
         */
        XLTableColumn appendColumn(std::string_view name);

        /**
         * @brief Get a column by its name
         * @param name The name of the column
         * @return The XLTableColumn object
         */
        XLTableColumn column(std::string_view name) const;

        /**
         * @brief Get a column by its 1-based ID
         * @param id The 1-based ID of the column
         * @return The XLTableColumn object
         */
        XLTableColumn column(uint32_t id) const;

        /**
         * @brief Get the auto filter object for the table.
         * @return XLAutoFilter The auto filter object.
         */
        XLAutoFilter autoFilter() const;

        /**
         * @brief Automatically resizes the table to fit the contiguous data in the worksheet.
         * The resize begins from the top-left corner of the current table's reference range.
         * It expands downwards and rightwards until it hits the boundaries of the used data.
         * @param worksheet The worksheet containing this table.
         */
        void resizeToFitData(const XLWorksheet& worksheet);

        /**
         * @brief Print the XML contents of this XLTables instance using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;
    };
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLTABLES_HPP
