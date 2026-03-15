#ifndef OPENXLSX_XLZIPARCHIVE_HPP
#define OPENXLSX_XLZIPARCHIVE_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"

#include <memory>
#include <string>

namespace OpenXLSX
{
    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLZipArchive
    {
    public:
        /**
         * @brief
         */
        XLZipArchive();

        /**
         * @brief
         * @param other
         */
        XLZipArchive(const XLZipArchive& other) = default;

        /**
         * @brief
         * @param other
         */
        XLZipArchive(XLZipArchive&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLZipArchive();

        /**
         * @brief
         * @param other
         * @return
         */
        XLZipArchive& operator=(const XLZipArchive& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLZipArchive& operator=(XLZipArchive&& other) noexcept = default;

        /**
         * @brief
         * @return
         */
        explicit operator bool() const;

        bool isValid() const;

        /**
         * @brief
         * @return
         */
        bool isOpen() const;

        /**
         * @brief
         * @param fileName
         */
        void open(const std::string& fileName);

        /**
         * @brief
         */
        void close();

        /**
         * @brief
         * @param path
         */
        void save(const std::string& path = "");

        /**
         * @brief
         * @param name
         * @param data
         */
        void addEntry(const std::string& name, const std::string& data);

        /**
         * @brief
         * @param entryName
         */
        void deleteEntry(const std::string& entryName);

        /**
         * @brief
         * @param name
         * @return
         */
        std::string getEntry(const std::string& name) const;

        /**
         * @brief
         * @param entryName
         * @return
         */
        bool hasEntry(const std::string& entryName) const;

        /**
         * @brief Get a list of all entry names in the archive.
         * @return A vector of strings containing the names of all entries.
         */
        std::vector<std::string> entryNames() const;

    private:
        struct LibZipApp;
        std::shared_ptr<LibZipApp> m_archive; /**< */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLZIPARCHIVE_HPP
