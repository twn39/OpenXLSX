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

        [[nodiscard]] bool isValid() const;

        /**
         * @brief
         * @return
         */
        [[nodiscard]] bool isOpen() const;

        /**
         * @brief
         * @param fileName
         */
        void open(std::string_view fileName);

        /**
         * @brief
         */
        void close();

        /**
         * @brief
         * @param path
         */
        void save(std::string_view path = "");

        /**
         * @brief
         * @param name
         * @param data
         */
        void addEntry(std::string_view name, std::string_view data);

        /**
         * @brief
         * @param entryName
         */
        void deleteEntry(std::string_view entryName);

        /**
         * @brief
         * @param name
         * @return
         */
        [[nodiscard]] std::string getEntry(std::string_view name) const;

        /**
         * @brief
         * @param entryName
         * @return
         */
        [[nodiscard]] bool hasEntry(std::string_view entryName) const;

        /**
         * @brief Get a list of all entry names in the archive.
         * @return A vector of strings containing the names of all entries.
         */
        [[nodiscard]] std::vector<std::string> entryNames() const;

    private:
        struct LibZipApp;
        std::shared_ptr<LibZipApp> m_archive; /**< */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLZIPARCHIVE_HPP
