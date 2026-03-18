#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

int main()
{
    try {
        XLDocument doc;
        std::string filename = "TableFeaturesDemo.xlsx";
        doc.create(filename, XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        wks.setName("Sales Data");

        // =========================================================================
        // Table 1: Employee Sales (Showcasing calculated columns & totals formulas)
        // =========================================================================
        
        // Write headers
        wks.cell("B2").value() = "EmpID";
        wks.cell("C2").value() = "Name";
        wks.cell("D2").value() = "Price";
        wks.cell("E2").value() = "Qty";
        wks.cell("F2").value() = "TotalSales"; // Calculated Column
        
        // Write data
        wks.cell("B3").value() = 1001; wks.cell("C3").value() = "Alice"; wks.cell("D3").value() = 150.5; wks.cell("E3").value() = 10;
        wks.cell("B4").value() = 1002; wks.cell("C4").value() = "Bob";   wks.cell("D4").value() = 200.0; wks.cell("E4").value() = 5;
        wks.cell("B5").value() = 1003; wks.cell("C5").value() = "Carol"; wks.cell("D5").value() = 99.9;  wks.cell("E5").value() = 20;
        
        // Create table and auto-initialize columns from the header row
        auto& tables = wks.tables();
        auto table1 = tables.add("SalesTable", "B2:F6"); // F6 is for the Totals Row
        table1.createColumnsFromRange(wks);
        
        // Configure Table Metadata and Style
        table1.setDisplayName("EmployeeSalesData");
        table1.setComment("This table calculates total sales per employee.");
        table1.setStyleName("TableStyleMedium9");
        table1.setShowRowStripes(true);
        table1.setShowTotalsRow(true);

        // Configure Calculated Column (TotalSales = Price * Qty)
        auto colTotal = table1.column("TotalSales");
        colTotal.setCalculatedColumnFormula("[@Price]*[@Qty]");
        
        // Write formulas into the sheet so Excel shows them immediately
        // In a real table, Excel auto-fills these if user types, but generating a file requires putting the formulas in cells
        wks.cell("F3").formula() = "SalesTable[@Price]*SalesTable[@Qty]";
        wks.cell("F4").formula() = "SalesTable[@Price]*SalesTable[@Qty]";
        wks.cell("F5").formula() = "SalesTable[@Price]*SalesTable[@Qty]";

        // Configure Totals Row Formulas
        auto colName = table1.column("Name");
        colName.setTotalsRowLabel("Grand Total:");
        wks.cell("C6").value() = "Grand Total:";
        
        auto colQty = table1.column("Qty");
        colQty.setTotalsRowFunction(XLTotalsRowFunction::Sum);
        wks.cell("E6").formula() = "SUBTOTAL(109,SalesTable[Qty])";
        
        colTotal.setTotalsRowFormula("SUM([TotalSales])");
        wks.cell("F6").formula() = "SUBTOTAL(109,SalesTable[TotalSales])";


        // =========================================================================
        // Table 2: Target Goals (Showcasing multiple tables in one sheet)
        // =========================================================================
        
        wks.cell("H2").value() = "Region";
        wks.cell("I2").value() = "Target";
        
        wks.cell("H3").value() = "North"; wks.cell("I3").value() = 5000;
        wks.cell("H4").value() = "South"; wks.cell("I4").value() = 8000;
        
        auto table2 = tables.add("RegionTargets", "H2:I4");
        table2.createColumnsFromRange(wks);
        table2.setStyleName("TableStyleLight1");
        
        doc.save();
        doc.close();
        
        std::cout << "Successfully generated " << filename << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        return 1;
    }
    return 0;
}
