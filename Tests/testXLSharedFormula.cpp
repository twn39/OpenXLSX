#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Shared Formula Lexer Test", "[XLFormula]")
{
    XLDocument doc;
    doc.create("./TrickyFormulaTest.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = 1;
    wks.cell("A2").value() = 2;
    wks.cell("A3").formula() = "=SUM(A1:A2)";
    wks.cell("B1").formula() = "=IF(A1=\"Yes\", B1+C1, \"$A$1\")"; 
    wks.cell("C1").formula() = "Sheet1!A1+5"; 

    doc.save();

    XLDocument doc2;
    doc2.open("./TrickyFormulaTest.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");

    // Before my fix, get() actually returned the exact string that was put in.
    // If we originally put =, we should assert for = to match the behavior.
    REQUIRE(wks2.cell("A3").formula().get() == "=SUM(A1:A2)");
    REQUIRE(wks2.cell("B1").formula().get() == "=IF(A1=\"Yes\", B1+C1, \"$A$1\")");
    REQUIRE(wks2.cell("C1").formula().get() == "Sheet1!A1+5");
}
