#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("TableFeatureDemo.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Write header
    wks.cell("B2").value() = "Employee ID";
    wks.cell("C2").value() = "Name";
    wks.cell("D2").value() = "Sales";
    wks.cell("E2").value() = "Bonus";

    // Write data
    wks.cell("B3").value() = 1001;
    wks.cell("C3").value() = "Alice";
    wks.cell("D3").value() = 5000;
    wks.cell("E3").value() = 500;

    wks.cell("B4").value() = 1002;
    wks.cell("C4").value() = "Bob";
    wks.cell("D4").value() = 4200;
    wks.cell("E4").value() = 420;

    wks.cell("B5").value() = 1003;
    wks.cell("C5").value() = "Charlie";
    wks.cell("D5").value() = 7500;
    wks.cell("E5").value() = 750;

    // Create table (starts from B2, just initial small ref)
    auto table = wks.tables().add("SalesTable", "B2:C3"); // Intentionally small, will be resized

    // Automatically expand table to fit the newly written data
    table.resizeToFitData(wks); // This should expand it to B2:E5

    // Add visual and totals features
    table.setStyleName("TableStyleMedium9");
    table.setShowRowStripes(true);
    table.setShowFirstColumn(true); // Highlight ID column
    table.setShowTotalsRow(true);

    // Setup columns and totals functions
    auto col1 = table.appendColumn("Employee ID");
    col1.setTotalsRowLabel("Grand Total:");

    auto col2 = table.appendColumn("Name");
    col2.setTotalsRowFunction(XLTotalsRowFunction::Count); // Count employees

    auto col3 = table.appendColumn("Sales");
    col3.setTotalsRowFunction(XLTotalsRowFunction::Sum); // Sum of sales

    auto col4 = table.appendColumn("Bonus");
    col4.setTotalsRowFunction(XLTotalsRowFunction::Average); // Average bonus

    // Set AutoFilter Custom filter
    auto filter = table.autoFilter();
    // Filter Sales (Col 3, which is 0-indexed as colId 2) > 4500
    filter.filterColumn(2).setCustomFilter("greaterThan", "4500");

    doc.save();
    doc.close();

    std::cout << "TableFeatureDemo.xlsx generated successfully." << std::endl;
    return 0;
}
