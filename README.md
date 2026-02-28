# OpenXLSX

OpenXLSX is a high-performance C++ library for reading, writing, creating, and modifying Microsoft Excel® files in the `.xlsx` format. It is designed to be fast, cross-platform, and has minimal external dependencies.

## 🚀 Key Features

- **Modern C++**: Built with C++17, ensuring type safety and modern abstractions.
- **High Performance**: Optimized for speed, capable of processing millions of cells per second.
- **Comprehensive Image Support**: (NEW) Insert and read images (PNG/JPEG) with automatic dimension detection and aspect ratio preservation.
- **Rich Formatting**: Support for styles, fonts, fills, borders, and conditional formatting.
- **Worksheet Management**: Create, clone, rename, and delete worksheets with validation.
- **Data Integrity**: Enforces OOXML standards, including strict XML declarations and metadata handling for seamless Excel compatibility.
- **Advanced Iterators**: Efficient cell and row iterators that minimize XML overhead.
- **Modern Testing**: Integrated with the latest **Catch2 v3.13.0** for robust verification.

## 🛠 Quick Start

```cpp
#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("Example.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    // Writing data
    wks.cell("A1").value() = "Hello, OpenXLSX!";
    wks.cell("B1").value() = 42;

    // (NEW) Inserting an image at B2 with 50% scaling
    // auto& drawing = wks.drawing();
    // std::string imagePath = doc.addImage("my_image.png", raw_binary_data);
    // auto rel = drawing.relationships().addRelationship(XLRelationshipType::Image, "../media/my_image.png");
    // drawing.addScaledImage(rel.id(), "my_image.png", "Description", raw_binary_data, 1, 1, 0.5);

    doc.save();
    return 0;
}
```

## 📦 Installation & Build

OpenXLSX uses CMake (3.15+) as its primary build system.

### Integration (FetchContent)
The easiest way to use OpenXLSX is via CMake's `FetchContent`:

```cmake
include(FetchContent)
FetchContent_Declare(
  OpenXLSX
  GIT_REPOSITORY https://github.com/troldal/OpenXLSX.git
  GIT_TAG        master # Or a specific commit
)
FetchContent_MakeAvailable(OpenXLSX)
target_link_libraries(my_project PRIVATE OpenXLSX::OpenXLSX)
```

### Manual Build
```bash
mkdir build && cd build
cmake ..
cmake --build . --config Release
```

## 🧪 Testing
The library features a comprehensive test suite powered by **Catch2 v3.13.0**.
To run the tests:
```bash
./build/output/OpenXLSXTests
```

## ⚠️ Development Conventions

### Unicode / UTF-8
**All string input and output must be in UTF-8 encoding.** OpenXLSX does not perform internal encoding conversion. Ensure your source files are saved in UTF-8 to avoid issues with hardcoded literals.

### Indexing
- **Sheet Indexing**: 1-based (consistent with Excel).
- **Row/Column Indexing**: Generally 1-based where it follows Excel conventions.

### Performance & Memory
OpenXLSX uses a DOM-based XML parser (PugiXML) for speed. For extremely large spreadsheets, memory consumption can be significant. A "compact mode" is available via CMake (`ENABLE_COMPACT_MODE`) to reduce memory footprint at the cost of some performance.

## 🤝 Credits
- [PugiXML](https://pugixml.org/) - Fast XML parsing.
- [miniz](https://github.com/richgel999/miniz) - ZIP compression/decompression.
- [Zippy](https://github.com/troldal/Zippy) - C++ wrapper for miniz.
- [Boost.Nowide](https://github.com/boostorg/nowide) - UTF-8 file system support on Windows.

---

<details>
<summary><b>Click to view Detailed Change Log</b></summary>

### 2026-02-28: Major Update
- **Image Support**: Implemented `XLDrawing` and `XLDrawingItem` for modern Spreadsheet Drawing ML.
- **Dimension Detection**: Added automatic PNG/JPEG dimension parsing in `XLUtilities.hpp`.
- **Scaling API**: Added `addScaledImage` to maintain original aspect ratios during insertion.
- **Catch2 v3**: Upgraded test suite to version 3.13.0 using modern CMake integration.
- **Excel Compatibility**: Fixed metadata persistence bugs in `[Content_Types].xml` and `.rels` files.
- **Refactoring**: Standardized `XLXmlFile` data access and reorganized utility headers.

</details>
