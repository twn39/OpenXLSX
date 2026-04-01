#include <OpenXLSX.hpp>
#include <iostream>
#include <vector>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("OpenXLSX_StreamStyles_Demo.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    wks.setName("Streaming Report");

    // 1. Setup global styles before streaming starts
    
    // Header Style: Dark Blue background, White Bold text, Centered
    XLStyle headerStyleDef;
    headerStyleDef.font.bold = true;
    headerStyleDef.font.color = XLColor("FFFFFF");
    headerStyleDef.fill.pattern = XLPatternSolid;
    headerStyleDef.fill.fgColor = XLColor("002060");
    headerStyleDef.alignment.horizontal = XLAlignCenter;
    auto headerStyle = doc.styles().findOrCreateStyle(headerStyleDef);

    // Error Style: Red text, Yellow background
    XLStyle errorStyleDef;
    errorStyleDef.font.color = XLColor("FF0000");
    errorStyleDef.font.bold = true;
    errorStyleDef.fill.pattern = XLPatternSolid;
    errorStyleDef.fill.fgColor = XLColor("FFFF00");
    auto errorStyle = doc.styles().findOrCreateStyle(errorStyleDef);

    // Currency Style: Number format for dollars
    XLStyle currencyStyleDef;
    currencyStyleDef.numberFormat = "$#,##0.00";
    auto currencyStyle = doc.styles().findOrCreateStyle(currencyStyleDef);

    // 2. Open Stream Writer
    auto stream = wks.streamWriter();

    // 3. Write Header Row (All styled)
    stream.appendRow({
        XLStreamCell("ID", headerStyle),
        XLStreamCell("Product Name", headerStyle),
        XLStreamCell("Price", headerStyle),
        XLStreamCell("Status", headerStyle)
    });

    // 4. Stream 100,000 rows of mock data with conditional styling
    std::cout << "Streaming 100,000 styled rows..." << std::endl;
    for (int i = 1; i <= 100000; ++i) {
        
        // Dynamically apply error style to out-of-stock items (every 10th item)
        bool outOfStock = (i % 10 == 0);
        
        std::vector<XLStreamCell> row;
        row.reserve(4);
        
        row.emplace_back(i); // Unstyled ID
        row.emplace_back("Enterprise Server Blade v" + std::to_string(i));
        row.emplace_back(2999.99 + (i % 100), currencyStyle); // Apply Currency Style
        
        if (outOfStock) {
            row.emplace_back("OUT OF STOCK", errorStyle); // Apply Error Style
        } else {
            row.emplace_back("In Stock"); // Unstyled
        }
        
        stream.appendRow(row);
    }

    // 5. Close and Save
    stream.close();
    doc.save();
    doc.close();

    std::cout << "Successfully generated OpenXLSX_StreamStyles_Demo.xlsx in the current directory!" << std::endl;
    return 0;
}