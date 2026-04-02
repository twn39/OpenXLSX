#include "XLThreadedComments.hpp"

using namespace OpenXLSX;

// ===== XLThreadedComment Implementation ===== //

XLThreadedComment::XLThreadedComment(XMLNode node) : m_node(node) {}

bool XLThreadedComment::valid() const { return m_node != nullptr; }

std::string XLThreadedComment::ref() const
{
    return m_node.attribute("ref").value();
}

std::string XLThreadedComment::id() const
{
    return m_node.attribute("id").value();
}

std::string XLThreadedComment::parentId() const
{
    return m_node.attribute("parentId").value();
}

std::string XLThreadedComment::personId() const
{
    return m_node.attribute("personId").value();
}

std::string XLThreadedComment::text() const
{
    XMLNode textNode = m_node.child("text");
    if (!textNode.empty()) return textNode.text().get();
    return "";
}

// ===== XLThreadedComments Implementation ===== //

XLThreadedComment XLThreadedComments::comment(const std::string& ref) const
{
    XMLNode root = xmlDocument().document_element();
    for (XMLNode node = root.first_child(); node; node = node.next_sibling()) {
        if (std::string(node.name()) == "threadedComment" && std::string(node.attribute("ref").value()) == ref) {
            // A top-level comment has no parentId
            if (node.attribute("parentId").empty()) {
                return XLThreadedComment(node);
            }
        }
    }
    return XLThreadedComment(XMLNode());
}

std::vector<XLThreadedComment> XLThreadedComments::replies(const std::string& parentId) const
{
    std::vector<XLThreadedComment> result;
    XMLNode root = xmlDocument().document_element();
    for (XMLNode node = root.first_child(); node; node = node.next_sibling()) {
        if (std::string(node.name()) == "threadedComment" && std::string(node.attribute("parentId").value()) == parentId) {
            result.emplace_back(node);
        }
    }
    return result;
}

// ===== XLPerson Implementation ===== //

XLPerson::XLPerson(XMLNode node) : m_node(node) {}

bool XLPerson::valid() const { return m_node != nullptr; }

std::string XLPerson::id() const
{
    return m_node.attribute("id").value();
}

std::string XLPerson::displayName() const
{
    return m_node.attribute("displayName").value();
}

// ===== XLPersons Implementation ===== //

XLPerson XLPersons::person(const std::string& id) const
{
    XMLNode root = xmlDocument().document_element();
    for (XMLNode node = root.first_child(); node; node = node.next_sibling()) {
        if (std::string(node.name()) == "person" && std::string(node.attribute("id").value()) == id) {
            return XLPerson(node);
        }
    }
    return XLPerson(XMLNode());
}
