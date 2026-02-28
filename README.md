# OpenXLSX

[![OpenXLSX CI](https://github.com/twn39/OpenXLSX/actions/workflows/ci.yml/badge.svg)](https://github.com/twn39/OpenXLSX/actions/workflows/ci.yml)

OpenXLSX is a high-performance C++ library for reading, writing, creating, and modifying Microsoft Excel® files in the `.xlsx` format. It is designed to be fast, cross-platform, and has minimal external dependencies.

## 🚀 Key Features

- **Modern C++**: Built with C++17, ensuring type safety and modern abstractions.
- **High Performance**: Optimized for speed, capable of processing millions of cells per second.
- **LTO Support**: Optional Link-Time Optimization (LTO/IPO) for maximum execution performance.
- **Comprehensive Image Support**: Insert and read images (PNG/JPEG) with automatic dimension detection and aspect ratio preservation.
- **Rich Formatting**: Support for styles, fonts, fills, borders, and conditional formatting.
- **Worksheet Management**: Create, clone, rename, and delete worksheets with validation.
- **Data Integrity**: Enforces OOXML standards, including strict XML declarations, core/app properties, and **Custom Properties**.
- **Memory Safety**: Integrated with **Microsoft GSL v4.2.1**, utilizing `span`, `not_null`, and `narrow` to ensure robust memory and type safety.
- **High-Performance ZIP**: Powered by **minizip-ng** and **zlib-ng**, featuring SIMD optimizations and full Zip64 support for massive datasets.
- **Advanced Iterators**: Efficient cell and row iterators that minimize XML overhead.
- **Modern Testing**: Integrated with **Catch2 v3.13.0** for robust verification.
- **Simplified Build**: A unified single-root CMake configuration for easy integration.

## 🛠 Quick Start

```cpp
#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("Example.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Writing data
    wks.cell("A1").value() = "Hello, OpenXLSX!";
    wks.cell("B1").value() = 42;

    doc.save();
    return 0;
}
```

## 📦 Installation & Build

OpenXLSX uses a simplified CMake (3.15+) build system.

### Integration (FetchContent)
The recommended way to use OpenXLSX is via CMake's `FetchContent`:

```cmake
include(FetchContent)
FetchContent_Declare(
  OpenXLSX
  GIT_REPOSITORY https://github.com/twn39/OpenXLSX.git
  GIT_TAG        master # Or a specific tag
)
FetchContent_MakeAvailable(OpenXLSX)
target_link_libraries(my_project PRIVATE OpenXLSX::OpenXLSX)
```

### Manual Build
To build the library and tests locally:
```bash
mkdir build && cd build
cmake .. -DCMAKE_BUILD_TYPE=Release
cmake --build . -j8
```

## 🧪 Testing
The library features a comprehensive test suite. To run the tests after building:
```bash
# From the build directory
./OpenXLSXTests
```

## ⚠️ Development Conventions

### Unicode / UTF-8
**All string input and output must be in UTF-8 encoding.** OpenXLSX does not perform internal encoding conversion. Ensure your source files are saved in UTF-8.

### Indexing
- **Sheet Indexing**: 1-based (consistent with Excel).
- **Row/Column Indexing**: Generally 1-based where it follows Excel conventions.

### Performance & Optimizations
The build system includes platform-specific optimizations for `Release` builds (e.g., `/O2` on MSVC, `-O3` on GCC/Clang) and supports **LTO (Link-Time Optimization)** which can be toggled via `OPENXLSX_ENABLE_LTO`.

## 🤝 Credits
- [PugiXML](https://pugixml.org/) - Fast XML parsing.
- [minizip-ng](https://github.com/zlib-ng/minizip-ng) - Next-generation ZIP manipulation.
- [zlib-ng](https://github.com/zlib-ng/zlib-ng) - SIMD-optimized compression backend.
- [Microsoft GSL](https://github.com/microsoft/GSL) - Guidelines Support Library for C++ safety.
- [Boost.Nowide](https://github.com/boostorg/nowide) - UTF-8 file system support on Windows.

---

<details>
<summary><b>Detailed Change Log (Feb 2026)</b></summary>

### 2026-02-28: Major Security & Performance Overhaul
- **Safety Hardening**: Fully integrated **Microsoft GSL v4.2.1**. Refactored XML parsing, cell addressing, and shared string management to use `gsl::span`, `gsl::not_null`, and `gsl::narrow`, eliminating potential buffer overflows and null pointer dereferences.
- **Modern ZIP Engine**: Migrated from `zippy/miniz` to **minizip-ng + zlib-ng**. Implemented a high-performance Pimpl-based `XLZipArchive` with memory-mapped caching and SIMD-accelerated compression.
- **Custom Metadata**: Added full support for **Custom Properties** (`custom.xml`), enabling users to read, write, and delete arbitrary document metadata tags.
- **Optimization Suite**: Switched core data structures to `std::unordered_map` for O(1) lookups in shared strings and zip entries.
- **Unified Build System**: Consolidated all sub-module CMake configurations into a single root `CMakeLists.txt`.
- **Enhanced Testing**: Reached 933 assertions with 100% pass rate, covering new safety contracts and metadata features.

</details>
