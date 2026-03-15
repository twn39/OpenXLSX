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

namespace OpenXLSX
{
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
         * @brief Print the XML contents of this XLTables instance using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;
    };
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLTABLES_HPP
