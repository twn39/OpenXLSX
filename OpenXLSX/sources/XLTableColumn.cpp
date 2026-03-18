#include "XLTableColumn.hpp"
#include <string>

namespace OpenXLSX {

    XLTableColumn::XLTableColumn(XMLNode node) : m_node(node) {}

    XLTableColumn::operator bool() const {
        return m_node != nullptr;
    }

    uint32_t XLTableColumn::id() const {
        if (!m_node) return 0;
        return m_node.attribute("id").as_uint();
    }

    std::string XLTableColumn::name() const {
        if (!m_node) return "";
        return m_node.attribute("name").value();
    }

    void XLTableColumn::setName(const std::string& name) {
        if (!m_node) return;
        auto attr = m_node.attribute("name");
        if (!attr) {
            m_node.append_attribute("name") = name.c_str();
        } else {
            attr.set_value(name.c_str());
        }
    }

    XLTotalsRowFunction XLTableColumn::totalsRowFunction() const {
        if (!m_node) return XLTotalsRowFunction::None;
        
        std::string funcStr = m_node.attribute("totalsRowFunction").value();
        if (funcStr == "sum") return XLTotalsRowFunction::Sum;
        if (funcStr == "min") return XLTotalsRowFunction::Min;
        if (funcStr == "max") return XLTotalsRowFunction::Max;
        if (funcStr == "average") return XLTotalsRowFunction::Average;
        if (funcStr == "count") return XLTotalsRowFunction::Count;
        if (funcStr == "countNums") return XLTotalsRowFunction::CountNums;
        if (funcStr == "stdDev") return XLTotalsRowFunction::StdDev;
        if (funcStr == "var") return XLTotalsRowFunction::Var;
        if (funcStr == "custom") return XLTotalsRowFunction::Custom;
        
        return XLTotalsRowFunction::None;
    }

    void XLTableColumn::setTotalsRowFunction(XLTotalsRowFunction func) {
        if (!m_node) return;

        std::string funcStr = "";
        switch (func) {
            case XLTotalsRowFunction::Sum: funcStr = "sum"; break;
            case XLTotalsRowFunction::Min: funcStr = "min"; break;
            case XLTotalsRowFunction::Max: funcStr = "max"; break;
            case XLTotalsRowFunction::Average: funcStr = "average"; break;
            case XLTotalsRowFunction::Count: funcStr = "count"; break;
            case XLTotalsRowFunction::CountNums: funcStr = "countNums"; break;
            case XLTotalsRowFunction::StdDev: funcStr = "stdDev"; break;
            case XLTotalsRowFunction::Var: funcStr = "var"; break;
            case XLTotalsRowFunction::Custom: funcStr = "custom"; break;
            case XLTotalsRowFunction::None: break;
        }

        if (funcStr.empty()) {
            m_node.remove_attribute("totalsRowFunction");
        } else {
            auto attr = m_node.attribute("totalsRowFunction");
            if (!attr) {
                m_node.append_attribute("totalsRowFunction") = funcStr.c_str();
            } else {
                attr.set_value(funcStr.c_str());
            }
        }
    }

    std::string XLTableColumn::totalsRowLabel() const {
        if (!m_node) return "";
        return m_node.attribute("totalsRowLabel").value();
    }

    void XLTableColumn::setTotalsRowLabel(const std::string& label) {
        if (!m_node) return;
        
        if (label.empty()) {
            m_node.remove_attribute("totalsRowLabel");
        } else {
            auto attr = m_node.attribute("totalsRowLabel");
            if (!attr) {
                m_node.append_attribute("totalsRowLabel") = label.c_str();
            } else {
                attr.set_value(label.c_str());
            }
        }
    }

} // namespace OpenXLSX