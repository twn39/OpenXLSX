#include <OpenXLSX.hpp>
#include <iostream>
using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("Validator.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    wks.cell("B2").value() = "Employee ID";
    wks.cell("C2").value() = "Name";
    wks.cell("B3").value() = 1001;
    wks.cell("C3").value() = "Alice";

    auto table = wks.tables().add("Table1", "B2:C4");
    table.setDisplayName("Table1");
    // VERY IMPORTANT: Table requires a totals row cell space to be valid!
    // If B2:C3 is data, the totals row needs to be B4:C4.
    // If we only allocated B2:C3, Excel will say the ref bounds don't match the row counts.
    table.setShowTotalsRow(true);
    
    auto col1 = table.appendColumn("Employee ID");
    col1.setTotalsRowLabel("Total");
    auto col2 = table.appendColumn("Name");
    col2.setTotalsRowFunction(XLTotalsRowFunction::Count);

    doc.save();
    return 0;
}
