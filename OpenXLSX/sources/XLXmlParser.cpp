// ===== External Includes ===== //
#include <cstring>    // strlen, memcpy, strcpy
#include <pugixml.hpp>

// // ===== OpenXLSX Includes ===== //
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    bool NO_XML_NS = true;    // default: no XML namespaces
    /**
     * @details this function is meaningless when PUGI_AUGMENTED is not defined / used
     */
    bool enable_xml_namespaces()
    {
#ifdef PUGI_AUGMENTED
        NO_XML_NS = false;
        return true;
#else
        return false;
#endif
    }
    /**
     * @details this function is meaningless when PUGI_AUGMENTED is not defined / used
     */
    bool disable_xml_namespaces()
    {
#ifdef PUGI_AUGMENTED
        NO_XML_NS = true;
        return true;
#else
        return false;
#endif
    }

    /**
     * @details return the node name without a potentially(!) existing namespace
     */
    const pugi::char_t* XMLNode::name_without_namespace(const pugi::char_t* name_) const
    {
        if (NO_XML_NS) return name_;    // if node namespaces are not stripped: return immediately
        int pos = 0;
        while (name_[pos] and name_[pos] != ':') ++pos;    // find namespace delimiter
        if (!name_[pos]) return name_;                    // if no delimiter found: return unmodified name
        return name_ + pos + 1;                           // if delimiter was found: return the name minus the namespace
    }

    /**
     * @details for creation of node children: copy this node's namespace without thread safety, using
     *  a static character array to avoid smart pointer performance impact
     */
    const pugi::char_t* XMLNode::namespaced_name_char(const pugi::char_t* name_, bool force_ns) const
    {
        // ===== If node has no namespace: Early pass-through return
        if (!name_begin or force_ns) return name_;
        if (name_begin + strlen(name_) > XLMaxNamespacedNameLen) {
            using namespace std::literals::string_literals;
            throw XLException("OpenXLSX_xml_node::"s + __func__ + ": strlen of "s + name_ + " exceeds XLMaxNamespacedNameLen "s +
                              std::to_string(XLMaxNamespacedNameLen));
        }

        static pugi::char_t
            namespaced_name_[XLMaxNamespacedNameLen + 1];    // static memory allocation for concatenating node namespace and name_

        // ===== If node has a namespace: create a namespaced version of name_
        memcpy(namespaced_name_, xml_node::name(), name_begin);    // copy the node namespace
        strcpy(namespaced_name_ + name_begin, name_);              // concatenate the name_ including the terminating zero
        return namespaced_name_;
    }

    /**
     * @details for creation of node children: copy this node's namespace with thread safety, using
     *  smart pointer with a trade-off for ~15-20% performance impact (increased runtime)
     * @note // 2024-08-18: made lambda parameter unnamed to eliminate -Wunused-parameter
     */
    std::shared_ptr<pugi::char_t> XMLNode::namespaced_name_shared_ptr(const pugi::char_t* name_, bool force_ns) const
    {
        // ===== If node has no namespace: Early pass-through return with noop-deleter
        if (!name_begin or force_ns) return std::shared_ptr<pugi::char_t>(const_cast<pugi::char_t*>(name_), [](pugi::char_t*) {});

        // ===== If node has a namespace: allocate memory for concatenation and create a namespaced version of name_
        std::shared_ptr<pugi::char_t> namespaced_name_(new pugi::char_t[name_begin + strlen(name_) + 1],
                                                       std::default_delete<pugi::char_t[]>());
        memcpy(namespaced_name_.get(), xml_node::name(), name_begin);    // copy the node namespace
        strcpy(namespaced_name_.get() + name_begin, name_);              // concatenate the name_ with terminating zero
        return namespaced_name_;
    }

    /**
     * @details determine the first xml_node child whose xml_node_type matches type_
     * @date 2024-04-25
     */
    XMLNode XMLNode::first_child_of_type(pugi::xml_node_type type_) const
    {
        if (_root) {
            XMLNode x = first_child();
            XMLNode l = last_child();
            while (x != l and x.type() != type_) x = x.next_sibling();
            if (x.type() == type_) return XMLNode(x);
        }
        return XMLNode();    // if no node matching type_ was found: return an empty node
    }

    /**
     * @details determine the last xml_node child whose xml_node_type matches type_
     * @date 2024-04-25
     */
    XMLNode XMLNode::last_child_of_type(pugi::xml_node_type type_) const
    {
        if (_root) {
            XMLNode f = first_child();
            XMLNode x = last_child();
            while (x != f and x.type() != type_) x = x.previous_sibling();
            if (x.type() == type_) return XMLNode(x);
        }
        return XMLNode();    // if no node matching type_ was found: return an empty node
    }

    /**
     * @details determine amount of xml_node children child whose xml_node_type matches type_
     * @date 2024-04-28
     */
    size_t XMLNode::child_count_of_type(pugi::xml_node_type type_) const
    {
        size_t counter = 0;
        if (_root) {
            XMLNode c = first_child_of_type(type_);
            while (!c.empty()) {
                ++counter;
                c = c.next_sibling_of_type(type_);
            }
        }
        return counter;
    }

    /**
     * @details determine the next xml_node sibling whose xml_node_type matches type_
     * @date 2024-04-26
     */
    XMLNode XMLNode::next_sibling_of_type(pugi::xml_node_type type_) const
    {
        for (pugi::xml_node n = pugi::xml_node::next_sibling(); n; n = n.next_sibling()) {
            if (n.type() == type_) return XMLNode(n);
        }
        return XMLNode();    // if no node matching type_ was found: return an empty node
    }

    /**
     * @details determine the previous xml_node sibling whose xml_node_type matches type_
     * @date 2024-04-26
     */
    XMLNode XMLNode::previous_sibling_of_type(pugi::xml_node_type type_) const
    {
        for (pugi::xml_node n = pugi::xml_node::previous_sibling(); n; n = n.previous_sibling()) {
            if (n.type() == type_) return XMLNode(n);
        }
        return XMLNode();    // if no node matching type_ was found: return an empty node
    }

    /**
     * @details determine the next xml_node sibling whose name() matches name_ and xml_node_type matches type_
     * @date 2024-04-26
     */
    XMLNode XMLNode::next_sibling_of_type(const pugi::char_t* name_, pugi::xml_node_type type_) const
    {
        for (pugi::xml_node n = pugi::xml_node::next_sibling(); n; n = n.next_sibling()) {
            if (n.type() == type_ and std::strcmp(n.name(), name_) == 0) return XMLNode(n);
        }
        return XMLNode();    // if no node matching type_ was found: return an empty node
    }

    /**
     * @details determine the previous xml_node sibling whose name() matches name_ and xml_node_type matches type_
     * @date 2024-04-26
     */
    XMLNode XMLNode::previous_sibling_of_type(const pugi::char_t* name_, pugi::xml_node_type type_) const
    {
        for (pugi::xml_node n = pugi::xml_node::previous_sibling(); n; n = n.previous_sibling()) {
            if (n.type() == type_ and std::strcmp(n.name(), name_) == 0) return XMLNode(n);
        }
        return XMLNode();    // if no node matching type_ was found: return an empty node
    }

}    // namespace OpenXLSX
