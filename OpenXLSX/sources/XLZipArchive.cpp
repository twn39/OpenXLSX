/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2026, Curry Tang

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

// ===== External Includes ===== //
#ifdef _WIN32
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#include <windows.h>
#endif

#include <mz.h>
#include <mz_strm.h>
#include <mz_zip.h>
#include <mz_zip_rw.h>
#include <unordered_map>
#include <vector>
#include <cstring>
#include <gsl/gsl>

// ===== OpenXLSX Includes ===== //
#include "XLZipArchive.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

class XLZipArchive::Impl
{
public:
    Impl() = default;
    ~Impl() { close(); }

    void open(const std::string& fileName)
    {
        m_fileName = fileName;
        m_entries.clear();

        void* reader = mz_zip_reader_create();
        if (mz_zip_reader_open_file(reader, fileName.c_str()) == MZ_OK) {
            int32_t err = mz_zip_reader_goto_first_entry(reader);
            while (err == MZ_OK) {
                mz_zip_file* file_info = nullptr;
                if (mz_zip_reader_entry_get_info(reader, &file_info) == MZ_OK) {
                    std::string entryName = file_info->filename;
                    
                    if (mz_zip_reader_entry_open(reader) == MZ_OK) {
                        std::vector<char> buffer(gsl::narrow<size_t>(file_info->uncompressed_size));
                        mz_zip_reader_entry_read(reader, buffer.data(), gsl::narrow<int32_t>(buffer.size()));
                        m_entries[entryName] = std::string(buffer.begin(), buffer.end());
                        mz_zip_reader_entry_close(reader);
                    }
                }
                err = mz_zip_reader_goto_next_entry(reader);
            }
            mz_zip_reader_close(reader);
        }
        mz_zip_reader_delete(&reader);
        m_isOpen = true;
    }

    void close()
    {
        m_entries.clear();
        m_isOpen = false;
    }

    void save(const std::string& path)
    {
        std::string savePath = path.empty() ? m_fileName : path;
        void* writer = mz_zip_writer_create();
        
        if (mz_zip_writer_open_file(writer, savePath.c_str(), 0, 0) != MZ_OK) {
            mz_zip_writer_delete(&writer);
            throw XLInternalError("Failed to open zip file for writing: " + savePath);
        }

        for (const auto& [name, data] : m_entries) {
            mz_zip_file file_info;
            memset(&file_info, 0, sizeof(file_info));
            file_info.filename = name.c_str();
            file_info.compression_method = MZ_COMPRESS_METHOD_DEFLATE;
            file_info.flag = MZ_ZIP_FLAG_UTF8;

            if (mz_zip_writer_entry_open(writer, &file_info) != MZ_OK) {
                mz_zip_writer_close(writer);
                mz_zip_writer_delete(&writer);
                throw XLInternalError("Failed to open zip entry for writing: " + name);
            }
            mz_zip_writer_entry_write(writer, data.c_str(), gsl::narrow<int32_t>(data.size()));
            mz_zip_writer_entry_close(writer);
        }

        mz_zip_writer_close(writer);
        mz_zip_writer_delete(&writer);
    }

    void addEntry(const std::string& name, const std::string& data) { m_entries[name] = data; }
    void deleteEntry(const std::string& entryName) { m_entries.erase(entryName); }
    std::string getEntry(const std::string& name) const { return m_entries.at(name); }
    bool hasEntry(const std::string& entryName) const { return m_entries.find(entryName) != m_entries.end(); }
    bool isOpen() const { return m_isOpen; }

private:
    std::string m_fileName;
    bool m_isOpen = false;
    std::unordered_map<std::string, std::string> m_entries;
};

XLZipArchive::XLZipArchive() : m_archive(nullptr) {}
XLZipArchive::~XLZipArchive() = default;
XLZipArchive::operator bool() const { return isValid(); }
bool XLZipArchive::isValid() const { return m_archive != nullptr; }
bool XLZipArchive::isOpen() const { return m_archive && m_archive->isOpen(); }
void XLZipArchive::open(const std::string& fileName)
{
    m_archive = std::make_shared<Impl>();
    m_archive->open(fileName);
}
void XLZipArchive::close()
{
    if (m_archive) m_archive->close();
    m_archive.reset();
}
void XLZipArchive::save(const std::string& path) { if (m_archive) m_archive->save(path); }
void XLZipArchive::addEntry(const std::string& name, const std::string& data) { if (m_archive) m_archive->addEntry(name, data); }
void XLZipArchive::deleteEntry(const std::string& entryName) { if (m_archive) m_archive->deleteEntry(entryName); }
std::string XLZipArchive::getEntry(const std::string& name) const { return m_archive ? m_archive->getEntry(name) : ""; }
bool XLZipArchive::hasEntry(const std::string& entryName) const { return m_archive && m_archive->hasEntry(entryName); }
