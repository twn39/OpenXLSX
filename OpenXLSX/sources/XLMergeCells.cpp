// ===== External Includes ===== //
#include <algorithm>
#include <iostream>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCellReference.hpp"
#include "XLException.hpp"
#include "XLMergeCells.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

namespace
{
    /**
     * @details
     */
    XLMergeCells::XLRect parseRange(std::string_view reference)
    {
        using namespace std::literals::string_literals;

        size_t pos = reference.find_first_of(':');
        if (pos == std::string_view::npos or pos < 2 or pos + 2 >= reference.length())
            throw XLInputError("XLMergeCells: not a valid range reference: \""s + std::string(reference) + "\""s);

        XLCellReference refTL(std::string(reference.substr(0, pos)));
        XLCellReference refBR(std::string(reference.substr(pos + 1)));

        uint32_t refTopRow    = refTL.row();
        uint16_t refFirstCol  = refTL.column();
        uint32_t refBottomRow = refBR.row();
        uint16_t refLastCol   = refBR.column();

        if (refBottomRow < refTopRow or refLastCol < refFirstCol or (refBottomRow == refTopRow and refLastCol == refFirstCol))
            throw XLInputError("XLMergeCells: not a valid range reference: \""s + std::string(reference) + "\""s);

        return {refTopRow, refFirstCol, refBottomRow, refLastCol};
    }
}    // namespace

/**
 * @details
 */
XLMergeCells::XLMergeCells() = default;

/**
 * @details
 */
XLMergeCells::XLMergeCells(const XMLNode& rootNode, std::vector<std::string_view> const& nodeOrder)
    : m_rootNode(rootNode),
      m_nodeOrder(nodeOrder),
      m_mergeCellsNode()
{
    if (m_rootNode.empty()) throw XLInternalError("XLMergeCells constructor: can not construct with an empty XML root node");

    m_mergeCellsNode  = m_rootNode.child("mergeCells");
    XMLNode mergeNode = m_mergeCellsNode.first_child_of_type(pugi::node_element);
    while (not mergeNode.empty()) {
        bool invalidNode = true;

        if (std::string(mergeNode.name()) == "mergeCell") {
            std::string_view ref = mergeNode.attribute("ref").value();
            if (not ref.empty()) {
                try {
                    // Populate cache during initialization to avoid repeated XML parsing.
                    m_mergeCache.push_back({std::string(ref), parseRange(ref)});
                    invalidNode = false;
                }
                catch (const XLInputError&) {
                    // Skip invalid nodes.
                }
            }
        }

        XMLNode nextNode = mergeNode.next_sibling_of_type(pugi::node_element);

        if (invalidNode) {
            std::cerr << "XLMergeCells constructor: invalid child element, either name is not mergeCell or reference is invalid:"
                      << std::endl;
            mergeNode.print(std::cerr);
            if (not nextNode.empty()) {
                // Cleanup whitespace before removing node to keep XML clean.
                while (mergeNode.next_sibling() != nextNode) m_mergeCellsNode.remove_child(mergeNode.next_sibling());
            }
            m_mergeCellsNode.remove_child(mergeNode);
        }

        mergeNode = nextNode;
    }

    if (not m_mergeCache.empty()) {
        XMLAttribute attr = m_mergeCellsNode.attribute("count");
        if (attr.empty()) attr = m_mergeCellsNode.append_attribute("count");
        attr.set_value(static_cast<unsigned long long>(m_mergeCache.size()));
    }
    else
        deleteAll();
}

/**
 * @details
 */
bool XLMergeCells::valid() const { return (not m_rootNode.empty()); }

/**
 * @details
 */
XLMergeIndex XLMergeCells::findMerge(std::string_view reference) const
{
    const auto iter =
        std::find_if(m_mergeCache.begin(), m_mergeCache.end(), [&](const XLMergeEntry& entry) { return reference == entry.reference; });

    return iter == m_mergeCache.end() ? XLMergeNotFound : static_cast<XLMergeIndex>(std::distance(m_mergeCache.begin(), iter));
}

/**
 * @details
 */
bool XLMergeCells::mergeExists(std::string_view reference) const { return findMerge(reference) >= 0; }

/**
 * @details
 */
XLMergeIndex XLMergeCells::findMergeByCell(std::string_view cellRef) const
{ return findMergeByCell(XLCellReference(std::string(cellRef))); }

XLMergeIndex XLMergeCells::findMergeByCell(XLCellReference cellRef) const
{
    const uint32_t row = cellRef.row();
    const uint16_t col = cellRef.column();

    // Use pre-parsed numerical ranges for high performance.
    const auto iter = std::find_if(m_mergeCache.begin(), m_mergeCache.end(), [row, col](const XLMergeEntry& entry) {
        return entry.rect.contains(row, col);
    });

    return iter == m_mergeCache.end() ? XLMergeNotFound : static_cast<XLMergeIndex>(std::distance(m_mergeCache.begin(), iter));
}

/**
 * @details
 */
size_t XLMergeCells::count() const { return m_mergeCache.size(); }

/**
 * @details
 */
const char* XLMergeCells::merge(XLMergeIndex index) const
{
    if (index < 0 or static_cast<size_t>(index) >= m_mergeCache.size()) {
        using namespace std::literals::string_literals;
        throw XLInputError("XLMergeCells::merge: index "s + std::to_string(index) + " is out of range"s);
    }
    return m_mergeCache[static_cast<size_t>(index)].reference.c_str();
}

/**
 * @details
 */
XLMergeIndex XLMergeCells::appendMerge(const std::string& reference)
{
    using namespace std::literals::string_literals;

    if (m_mergeCache.size() >= XLMaxMergeCells)
        throw XLInputError("XLMergeCells::appendMerge: exceeded max merge cells count "s + std::to_string(XLMaxMergeCells));

    const XLRect newRect = parseRange(reference);

    // Guard against overlaps using numerical range comparison.
    for (const auto& entry : m_mergeCache) {
        if (entry.rect.overlaps(newRect))
            throw XLInputError("XLMergeCells::appendMerge: reference \""s + reference + "\" overlaps with existing reference \""s +
                               entry.reference + "\""s);
    }

    if (m_mergeCellsNode.empty())
        m_mergeCellsNode = appendAndGetNode(m_rootNode, "mergeCells", m_nodeOrder);

    XMLNode insertAfter = m_mergeCellsNode.last_child_of_type(pugi::node_element);
    XMLNode newMerge{};
    if (insertAfter.empty())
        newMerge = m_mergeCellsNode.prepend_child("mergeCell");
    else
        newMerge = m_mergeCellsNode.insert_child_after("mergeCell", insertAfter);

    if (newMerge.empty()) throw XLInternalError("XLMergeCells::appendMerge: failed to insert reference: \""s + reference + "\""s);
    newMerge.append_attribute("ref").set_value(reference.c_str());

    m_mergeCache.push_back({reference, newRect});

    XMLAttribute attr = m_mergeCellsNode.attribute("count");
    if (attr.empty()) attr = m_mergeCellsNode.append_attribute("count");
    attr.set_value(static_cast<unsigned long long>(m_mergeCache.size()));

    return static_cast<XLMergeIndex>(m_mergeCache.size() - 1);
}

/**
 * @details
 */
void XLMergeCells::deleteMerge(XLMergeIndex index)
{
    using namespace std::literals::string_literals;

    if (index < 0 or static_cast<size_t>(index) >= m_mergeCache.size())
        throw XLInputError("XLMergeCells::deleteMerge: index "s + std::to_string(index) + " is out of range"s);

    XLMergeIndex curIndex = 0;
    XMLNode      node     = m_mergeCellsNode.first_child_of_type(pugi::node_element);
    while (curIndex < index and not node.empty()) {
        node = node.next_sibling_of_type(pugi::node_element);
        ++curIndex;
    }

    if (node.empty()) throw XLInternalError("XLMergeCells::deleteMerge: mismatch between size of mergeCells XML node and internal cache");

    while (node.previous_sibling().type() == pugi::node_pcdata) m_mergeCellsNode.remove_child(node.previous_sibling());
    m_mergeCellsNode.remove_child(node);

    m_mergeCache.erase(m_mergeCache.begin() + index);

    if (not m_mergeCache.empty()) {
        XMLAttribute attr = m_mergeCellsNode.attribute("count");
        if (attr.empty()) attr = m_mergeCellsNode.append_attribute("count");
        attr.set_value(static_cast<unsigned long long>(m_mergeCache.size()));
    }
    else
        deleteAll();
}

/**
 * @details
 */
void XLMergeCells::deleteAll()
{
    m_mergeCache.clear();
    m_rootNode.remove_child(m_mergeCellsNode);
    m_mergeCellsNode = XMLNode();
}

/**
 * @details
 */
void XLMergeCells::print(std::basic_ostream<char>& ostr) const { m_mergeCellsNode.print(ostr); }
