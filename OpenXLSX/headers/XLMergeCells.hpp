#ifndef OPENXLSX_XLMERGECELLS_HPP
#define OPENXLSX_XLMERGECELLS_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

#include <deque>
#include <limits>
#include <memory>
#include <ostream>
#include <string>
#include <string_view>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCellReference.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    typedef int32_t              XLMergeIndex;
    constexpr const XLMergeIndex XLMergeNotFound = -1;

    constexpr size_t XLMaxMergeCells = (std::numeric_limits<XLMergeIndex>::max)();

    /**
     * @brief Manages merged cell ranges in a worksheet.
     * @details This class handles the <mergeCells> element. It maintains a numerical cache of ranges
     * to avoid expensive XML parsing and string manipulation during overlap checks and lookups.
     */
    class OPENXLSX_EXPORT XLMergeCells
    {
    public:
        /**
         * @brief Numerical representation of a cell range for fast overlap checks.
         */
        struct XLRect
        {
            uint32_t top;
            uint16_t left;
            uint32_t bottom;
            uint16_t right;

            [[nodiscard]] bool contains(uint32_t row, uint16_t col) const noexcept
            { return row >= top && row <= bottom && col >= left && col <= right; }

            [[nodiscard]] bool overlaps(const XLRect& other) const noexcept
            { return top <= other.bottom && bottom >= other.top && left <= other.right && right >= other.left; }
        };

        struct XLMergeEntry
        {
            std::string reference;
            XLRect      rect;
        };

    public:
        XLMergeCells();

        /**
         * @param rootNode The worksheet root node.
         * @param nodeOrder Required child sequence for proper OOXML schema compliance.
         */
        explicit XLMergeCells(const XMLNode& rootNode, std::vector<std::string_view> const& nodeOrder);

        ~XLMergeCells() = default;
        XLMergeCells(const XLMergeCells& other) = default;
        XLMergeCells(XLMergeCells&& other) noexcept = default;
        XLMergeCells& operator=(const XLMergeCells& other) = default;
        XLMergeCells& operator=(XLMergeCells&& other) noexcept = default;

        /**
         * @return true if initialized with a valid worksheet node.
         */
        [[nodiscard]] bool valid() const;

        /**
         * @return 0-based index or XLMergeNotFound.
         */
        [[nodiscard]] XLMergeIndex findMerge(std::string_view reference) const;

        [[nodiscard]] bool mergeExists(std::string_view reference) const;

        /**
         * @return The index of the merge range containing the given cell.
         */
        [[nodiscard]] XLMergeIndex findMergeByCell(std::string_view cellRef) const;
        [[nodiscard]] XLMergeIndex findMergeByCell(XLCellReference cellRef) const;

        [[nodiscard]] size_t count() const;

        /**
         * @return The range reference string (e.g., "A1:C3").
         */
        [[nodiscard]] const char* merge(XLMergeIndex index) const;

        [[nodiscard]] const char* operator[](XLMergeIndex index) const { return merge(index); }

        /**
         * @return The index of the newly appended merge.
         * @throws XLInputError If the range overlaps with an existing merge.
         */
        XLMergeIndex appendMerge(const std::string& reference);

        /**
         * @details Invalidation: previous indices are invalidated after deletion.
         */
        void deleteMerge(XLMergeIndex index);

        /**
         * @brief Removes all merged ranges and the <mergeCells> XML container.
         */
        void deleteAll();

        void print(std::basic_ostream<char>& ostr) const;

    private:
        XMLNode                       m_rootNode;
        std::vector<std::string_view> m_nodeOrder;
        XMLNode                       m_mergeCellsNode;
        std::deque<XLMergeEntry>      m_mergeCache; /**< Numerical cache to optimize lookups and overlap detection. */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLMERGECELLS_HPP
