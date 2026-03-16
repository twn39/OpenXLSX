#ifndef OPENXLSX_XLRELATIONSHIPS_HPP
#define OPENXLSX_XLRELATIONSHIPS_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <random>    // std::mt19937
#include <string>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    /**
     * @brief Enable use of random (relationship) IDs
     */
    void UseRandomIDs();

    /**
     * @brief Disable use of random (relationship) IDs (default behavior)
     */
    void UseSequentialIDs();

    /**
     * @brief Return a 32 bit random value
     * @return A 32 bit random value
     */
    extern std::mt19937 Rand32;

    /**
     * @brief Return a 64 bit random value (by invoking Rand32 twice)
     * @return A 64 bit random value
     */
    uint64_t Rand64();

    /**
     * @brief Initialize XLRand32 data source
     * @param pseudoRandom If true, sequence will be reproducible with a constant seed
     */
    void InitRandom(bool pseudoRandom = false);

    class XLRelationships;

    class XLRelationshipItem;

    /**
     * @brief An enum of the possible relationship (or XML document) types used in relationship (.rels) XML files.
     */
    enum class XLRelationshipType {
        CoreProperties,
        ExtendedProperties,
        CustomProperties,
        Workbook,
        Worksheet,
        Chartsheet,
        Dialogsheet,
        Macrosheet,
        CalculationChain,
        ExternalLink,
        ExternalLinkPath,
        Theme,
        Styles,
        Chart,
        ChartStyle,
        ChartColorStyle,
        Image,
        Drawing,
        VMLDrawing,
        SharedStrings,
        PrinterSettings,
        VBAProject,
        ControlProperties,
        Comments,
        Table,
        Hyperlink,
        Unknown
    };
}    //     namespace OpenXLSX

namespace OpenXLSX_XLRelationships
{    // special namespace to avoid naming conflict with another GetStringFromType function
    using namespace OpenXLSX;
    /**
     * @brief helper function, used only within module and from XLProperties.cpp / XLAppProperties::createFromTemplate
     * @param type the XLRelationshipType for which to return the correct XML string
     */
    std::string GetStringFromType(XLRelationshipType type);
}    //    namespace OpenXLSX_XLRelationships

namespace OpenXLSX
{
    /**
     * @brief An encapsulation of a relationship item, i.e. an XML file in the document, its type and an ID number.
     */
    class OPENXLSX_EXPORT XLRelationshipItem
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief Default constructor. Creates an empty/invalid relationship item.
         */
        XLRelationshipItem() = default;

        /**
         * @brief Constructor. New items should only be created through an XLRelationship object.
         * @param node An XMLNode object with the relationship item.
         */
        explicit XLRelationshipItem(const XMLNode& node);

        /**
         * @brief Rule of Zero: Default special member functions for value-based XMLNode storage.
         */
        ~XLRelationshipItem() = default;
        XLRelationshipItem(const XLRelationshipItem& other) = default;
        XLRelationshipItem(XLRelationshipItem&& other) noexcept = default;
        XLRelationshipItem& operator=(const XLRelationshipItem& other) = default;
        XLRelationshipItem& operator=(XLRelationshipItem&& other) noexcept = default;

        /**
         * @brief Get the type of the current relationship item.
         */
        [[nodiscard]] XLRelationshipType type() const;

        /**
         * @brief Get the target, i.e. the internal package path or external URI this relationship points to.
         */
        [[nodiscard]] std::string target() const;

        /**
         * @brief Get the unique relationship identifier (e.g., "rId1").
         */
        [[nodiscard]] std::string id() const;

        /**
         * @brief Determine if this item is empty or invalid.
         */
        [[nodiscard]] bool empty() const;

    private:                                 // ---------- Private Member Variables ---------- //
        XMLNode m_relationshipNode; /**< The underlying XML element for this relationship. Held by value to avoid heap allocations. */
    };

    // ================================================================================
    // XLRelationships Class
    // ================================================================================

    /**
     * @brief An encapsulation of relationship files (.rels files) in an Excel document package.
     */
    class OPENXLSX_EXPORT XLRelationships : public XLXmlFile
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLRelationships() = default;

        /**
         * @brief
         * @param xmlData
         * @param pathTo Initialize m_path from this: the path to the relationships file origin of xmlData
         * @note m_path is used to resolve relative relationship target paths to an absolute
         */
        explicit XLRelationships(XLXmlData* xmlData, std::string pathTo);

        /**
         * @brief Destructor
         */
        ~XLRelationships();

        /**
         * @brief
         * @param other
         */
        XLRelationships(const XLRelationships& other) = default;

        /**
         * @brief
         * @param other
         */
        XLRelationships(XLRelationships&& other) noexcept = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLRelationships& operator=(const XLRelationships& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLRelationships& operator=(XLRelationships&& other) noexcept = default;

        /**
         * @brief Look up a relationship item by ID.
         * @param id The unique identifier (e.g. "rId1").
         */
        [[nodiscard]] XLRelationshipItem relationshipById(std::string_view id) const;

        /**
         * @brief Look up a relationship item by Target path.
         * @param target The Target URI or package path.
         * @param throwIfNotFound If true, missing targets trigger an exception; otherwise returns an empty item.
         */
        [[nodiscard]] XLRelationshipItem relationshipByTarget(std::string_view target, bool throwIfNotFound = true) const;

        /**
         * @brief Get all relationship items in this file.
         */
        [[nodiscard]] std::vector<XLRelationshipItem> relationships() const;

        /**
         * @brief Remove a relationship by its ID.
         */
        void deleteRelationship(std::string_view relID);

        /**
         * @brief Remove a relationship by item reference.
         */
        void deleteRelationship(const XLRelationshipItem& item);

        /**
         * @brief Create and append a new relationship to the collection.
         * @param type OOXML relationship type (schema URI).
         * @param target Internal path or external URI.
         * @param isExternal Marking this relationship as 'External' avoids path normalization issues.
         */
        XLRelationshipItem addRelationship(XLRelationshipType type, std::string_view target, bool isExternal = false);

        /**
         * @brief Determine if a specific target path is already registered.
         */
        [[nodiscard]] bool targetExists(std::string_view target) const;

        /**
         * @brief Determine if a specific relationship ID is already in use.
         */
        [[nodiscard]] bool idExists(std::string_view id) const;

        /**
         * @brief print the XML contents of the relationships document using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;

        // ---------- Protected Member Functions ---------- //
    protected:
        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------
    private:
        std::string m_path;    // the path - within the XLSX file - to the relationships file on which this object is instantiated
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLRELATIONSHIPS_HPP
