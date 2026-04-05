#include <OpenXLSX.hpp>
#include <fstream>
#include <string>
#include <vector>
#include <cstdio>
#include <cstdlib>

using namespace OpenXLSX;

extern "C" int LLVMFuzzerTestOneInput(const uint8_t* data, size_t size)
{
    if (size == 0) return 0;

    std::string tempFilename = "/tmp/openxlsx_fuzz_" + std::to_string(rand()) + ".xlsx";

    std::ofstream out(tempFilename, std::ios::binary);
    if (!out.is_open()) return 0;
    out.write(reinterpret_cast<const char*>(data), size);
    out.close();

    try {
        XLDocument doc;
        doc.open(tempFilename);

        auto wbk = doc.workbook();
        for (const auto& sheetName : wbk.worksheetNames()) {
            auto wks = wbk.worksheet(sheetName);
            (void)wks.rowCount();
        }
        
        doc.close();
    }
    catch (const XLException&) {
    }
    catch (const std::exception&) {
    }

    std::remove(tempFilename.c_str());

    return 0;
}
