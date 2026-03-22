#ifndef OPENXLSX_XLCELLRANGE_HPP
#define OPENXLSX_XLCELLRANGE_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

// ===== External Includes ===== //
#include <memory>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCell.hpp"
#include "XLCellIterator.hpp"
#include "XLCellReference.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    /**
     * @brief Represents a rectangular area of cells within a worksheet.
     * @details This class provides a high-level interface for bulk operations on cells, 
     * such as clearing values, setting formats, or iterating over a subset of the worksheet.
     * It maintains a cache of column styles to ensure efficient cell creation during iteration.
     */
    class OPENXLSX_EXPORT XLCellRange
    {
        friend class XLCellIterator;

    public:
        /**
         * @brief Constructs an uninitialized range.
         */
        XLCellRange();

        /**
         * @param dataNode The XML node containing sheet data.
         * @param topLeft Top-left boundary of the range.
         * @param bottomRight Bottom-right boundary of the range.
         * @param sharedStrings Reference to the workbook's shared strings table.
         * @throws XLInputError if topLeft is not truly to the top-left of bottomRight.
         */
        explicit XLCellRange(const XMLNode&         dataNode,
                             const XLCellReference& topLeft,
                             const XLCellReference& bottomRight,
                             const XLSharedStrings& sharedStrings);

        ~XLCellRange() = default;
        XLCellRange(const XLCellRange& other) = default;
        XLCellRange(XLCellRange&& other) noexcept = default;
        XLCellRange& operator=(const XLCellRange& other) = default;
        XLCellRange& operator=(XLCellRange&& other) noexcept = default;

        /**
         * @brief Scans the worksheet for column-level styles and caches them.
         * @details This is called automatically during construction to ensure XLCellIterator 
         * can apply default styles to newly created cells without repeated XML lookups.
         */
        void fetchColumnStyles();

        [[nodiscard]] XLCellReference topLeft() const;
        [[nodiscard]] XLCellReference bottomRight() const;

        /**
         * @return The range reference string (e.g., "A1:C3"), or an empty string if uninitialized.
         */
        [[nodiscard]] std::string address() const;

        [[nodiscard]] uint32_t numRows() const;
        [[nodiscard]] uint16_t numColumns() const;

        [[nodiscard]] XLCellIterator begin() const;
        [[nodiscard]] XLCellIterator end() const;

        /**
         * @brief Returns true if the range is uninitialized or points to an invalid worksheet node.
         */
        [[nodiscard]] bool empty() const;

        /**
         * @brief Clears the values of all cells within the range.
         */
        void clear();

        /**
         * @brief Assigns a single value to every cell in the range.
         */
        template<typename T,
                 typename = std::enable_if_t<
                     std::is_integral_v<T> or std::is_floating_point_v<T> or std::is_same_v<std::decay_t<T>, std::string> ||
                     std::is_same_v<std::decay_t<T>, std::string_view> or std::is_same_v<std::decay_t<T>, const char*> ||
                     std::is_same_v<std::decay_t<T>, char*> or std::is_same_v<T, XLDateTime>>>
        XLCellRange& operator=(T value)
        {
            for (auto it = begin(); it != end(); ++it) it->value() = value;
            return *this;
        }

        /**
         * @brief Applies a cell format (style) to all cells in the range.
         * @param cellFormatIndex The index in the workbook's style sheet.
         * @return true on success.
         */
        XLCellRange& setFormat(XLStyleIndex cellFormatIndex);

        /**
         * @brief Calculates the intersection area between this range and another.
         * @return A new range representing the common area, or an empty range if they don't overlap.
         */
        [[nodiscard]] XLCellRange intersect(const XLCellRange& other) const;

    private:
        XMLNode                   m_dataNode;
        XLCellReference           m_topLeft;
        XLCellReference           m_bottomRight;
        XLSharedStringsRef        m_sharedStrings;
        std::vector<XLStyleIndex> m_columnStyles; /**< Cached column-level styles for fast cell creation. */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLCELLRANGE_HPP
