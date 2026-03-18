#pragma once

#include "XLXmlParser.hpp"
#include <string>

namespace OpenXLSX {

    /**
     * @brief Enum class defining the possible functions for a table's totals row.
     */
    enum class XLTotalsRowFunction {
        None,
        Sum,
        Min,
        Max,
        Average,
        Count,
        CountNums,
        StdDev,
        Var,
        Custom
    };

    /**
     * @brief The XLTableColumn class encapsulates the <tableColumn> XML node.
     * It allows setting different kinds of column properties in a table.
     */
    class XLTableColumn {
    public:
        /**
         * @brief Constructor
         * @param node The <tableColumn> XMLNode
         */
        explicit XLTableColumn(XMLNode node);

        /**
         * @brief Check if the TableColumn object is valid (has a corresponding XML node).
         * @return true if valid, false otherwise.
         */
        operator bool() const;

        /**
         * @brief Get the ID of the table column.
         * @return The column ID.
         */
        uint32_t id() const;

        /**
         * @brief Get the name of the table column.
         * @return The column name.
         */
        std::string name() const;

        /**
         * @brief Set the name of the table column.
         * @param name The new name of the column.
         */
        void setName(const std::string& name);

        /**
         * @brief Get the totals row function for this column.
         * @return The XLTotalsRowFunction.
         */
        XLTotalsRowFunction totalsRowFunction() const;

        /**
         * @brief Set the totals row function for this column.
         * @param func The XLTotalsRowFunction to set.
         */
        void setTotalsRowFunction(XLTotalsRowFunction func);

        /**
         * @brief Get the totals row label for this column.
         * @return The custom string label.
         */
        std::string totalsRowLabel() const;

        /**
         * @brief Set the totals row label for this column.
         * @param label The string to set as label.
         */
        void setTotalsRowLabel(const std::string& label);

    private:
        XMLNode m_node;
    };

} // namespace OpenXLSX
