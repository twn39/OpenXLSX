#ifndef OPENXLSX_XLPROPERTIES_HPP
#define OPENXLSX_XLPROPERTIES_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <string>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"

namespace OpenXLSX
{
    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLProperties : public XLXmlFile
    {
    private:
        /**
         * @brief constructor helper function: create core.xml content from template
         * @param workbook
         */
        void createFromTemplate();

        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:
        /**
         * @brief
         */
        XLProperties() = default;

        /**
         * @brief
         * @param xmlData
         */
        explicit XLProperties(XLXmlData* xmlData);

        /**
         * @brief
         * @param other
         */
        XLProperties(const XLProperties& other) = default;

        /**
         * @brief
         * @param other
         */
        XLProperties(XLProperties&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLProperties();

        /**
         * @brief
         * @param other
         * @return
         */
        XLProperties& operator=(const XLProperties& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLProperties& operator=(XLProperties&& other) = default;

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        void setProperty(const std::string& name, const std::string& value);

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        void setProperty(const std::string& name, int value);

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        void setProperty(const std::string& name, double value);

        /**
         * @brief
         * @param name
         * @return
         */
        std::string property(const std::string& name) const;

        /**
         * @brief
         * @param name
         */
        void deleteProperty(const std::string& name);

        //----------------------------------------------------------------------------------------------------------------------
        //           Protected Member Functions
        //----------------------------------------------------------------------------------------------------------------------
    };

    /**
     * @brief This class is a specialization of the XLAbstractXMLFile, with the purpose of the representing the
     * document app properties in the app.xml file (docProps folder) in the .xlsx package.
     */
    class OPENXLSX_EXPORT XLAppProperties : public XLXmlFile
    {
    private:
        /**
         * @brief constructor helper function: create app.xml content from template
         * @param workbook
         */
        void createFromTemplate(XMLDocument const& workbookXml);

        //--------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //--------------------------------------------------------------------------------------------------------------

    public:
        /**
         * @brief
         */
        XLAppProperties() = default;

        /**
         * @brief enable XLAppProperties to re-create a worksheet list in docProps/app.xml <TitlesOfParts> element from workbookXml
         * @param xmlData
         * @param workbook
         */
        explicit XLAppProperties(XLXmlData* xmlData, XMLDocument const& workbookXml);

        /**
         * @brief
         * @param xmlData
         */
        explicit XLAppProperties(XLXmlData* xmlData);

        /**
         * @brief
         * @param other
         */
        XLAppProperties(const XLAppProperties& other) = default;

        /**
         * @brief
         * @param other
         */
        XLAppProperties(XLAppProperties&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLAppProperties();

        /**
         * @brief
         * @param other
         * @return
         */
        XLAppProperties& operator=(const XLAppProperties& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLAppProperties& operator=(XLAppProperties&& other) noexcept = default;

        /**
         * @brief update the "HeadingPairs" entry for "Worksheets" *and* the "TitlesOfParts" vector size
         * @param increment change the sheet count by this (negative = decrement)
         * @throws XLInternalError when sheet count would become < 1
         */
        void incrementSheetCount(int16_t increment);

        /**
         * @brief initialize <TitlesOfParts> to contain all and only entries from workbookSheetNames & ensure HeadingPairs entry for
         * Worksheets has the correct count
         * @param workbookSheetNames the vector of sheet names as returned by XLWorkbook::sheetNames()
         * @throws XLInternalError thrown by the underlying sheetNames call upon failure
         */
        void alignWorksheets(std::vector<std::string> const& workbookSheetNames);

        /**
         * @brief
         * @param title
         * @return
         */
        void addSheetName(const std::string& title);

        /**
         * @brief
         * @param title
         */
        void deleteSheetName(const std::string& title);

        /**
         * @brief
         * @param oldTitle
         * @param newTitle
         */
        void setSheetName(const std::string& oldTitle, const std::string& newTitle);

        /**
         * @brief
         * @param name
         * @param value
         */
        void addHeadingPair(const std::string& name, int value);

        /**
         * @brief
         * @param name
         */
        void deleteHeadingPair(const std::string& name);

        /**
         * @brief
         * @param name
         * @param newValue
         */
        void setHeadingPair(const std::string& name, int newValue);

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        void setProperty(const std::string& name, const std::string& value);

        /**
         * @brief
         * @param name
         * @return
         */
        std::string property(const std::string& name) const;

        /**
         * @brief
         * @param name
         */
        void deleteProperty(const std::string& name);

        /**
         * @brief
         * @param sheetName
         * @return
         */
        void appendSheetName(const std::string& sheetName);

        /**
         * @brief
         * @param sheetName
         * @return
         */
        void prependSheetName(const std::string& sheetName);

        /**
         * @brief
         * @param sheetName
         * @param index
         * @return
         */
        void insertSheetName(const std::string& sheetName, unsigned int index);
    };

    /**
     * @brief The XLCustomProperties class encapsulates the functionality for reading and writing
     * custom document properties in docProps/custom.xml. Custom properties allow users to attach
     * arbitrary metadata to the workbook, which can be seen in Excel's file properties.
     */
    class OPENXLSX_EXPORT XLCustomProperties : public XLXmlFile
    {
    private:
        /**
         * @brief Initialize a new custom.xml file with the required namespaces if it doesn't exist.
         */
        void createFromTemplate();

    public:
        /**
         * @brief Default constructor.
         */
        XLCustomProperties() = default;

        /**
         * @brief Constructor that binds to existing XML data. If the XML is empty, it initializes it from template.
         * @param xmlData Pointer to the XML data managed by the parent document.
         */
        explicit XLCustomProperties(XLXmlData* xmlData);

        /**
         * @brief Copy constructor.
         */
        XLCustomProperties(const XLCustomProperties& other) = default;

        /**
         * @brief Move constructor.
         */
        XLCustomProperties(XLCustomProperties&& other) noexcept = default;

        /**
         * @brief Destructor.
         */
        ~XLCustomProperties();

        /**
         * @brief Copy assignment operator.
         */
        XLCustomProperties& operator=(const XLCustomProperties& other) = default;

        /**
         * @brief Move assignment operator.
         */
        XLCustomProperties& operator=(XLCustomProperties&& other) noexcept = default;

        /**
         * @brief Set a string custom property. If a property with the same name exists, it will be updated.
         * We use vt:lpwstr for strings to ensure compatibility with most Excel versions and support for Unicode.
         * @param name The name of the property.
         * @param value The string value.
         */
        void setProperty(const std::string& name, const std::string& value);

        /**
         * @brief Overload for string literals to prevent unwanted implicit conversion to bool.
         * @param name The name of the property.
         * @param value The string literal value.
         */
        void setProperty(const std::string& name, const char* value);

        /**
         * @brief Set an integer custom property using vt:i4 (32-bit signed integer).
         * @param name The name of the property.
         * @param value The integer value.
         */
        void setProperty(const std::string& name, int value);

        /**
         * @brief Set a double custom property using vt:r8 (64-bit floating point).
         * @param name The name of the property.
         * @param value The double value.
         */
        void setProperty(const std::string& name, double value);

        /**
         * @brief Set a boolean custom property using vt:bool.
         * @param name The name of the property.
         * @param value The boolean value.
         */
        void setProperty(const std::string& name, bool value);

        /**
         * @brief Retrieve the value of a custom property as a string, regardless of its underlying XML type.
         * Returns an empty string if the property is not found.
         * @param name The name of the property to find.
         * @return The property value as a string.
         */
        std::string property(const std::string& name) const;

        /**
         * @brief Remove a custom property by name. Does nothing if the property doesn't exist.
         * @param name The name of the property to delete.
         */
        void deleteProperty(const std::string& name);
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLPROPERTIES_HPP
