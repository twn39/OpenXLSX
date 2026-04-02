#include "XLCrypto.hpp"
#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"
#include <fstream>
#include <vector>

using namespace OpenXLSX;

TEST_CASE("Crypto Encryption and Decryption Cycle") {
    SECTION("Standard Encryption Export and Readback") {
        XLDocument doc;
        doc.create("Tests/Fixtures/OpenXLSX_Encrypted_Cycle_Test.xlsx", XLForceOverwrite);
        
        auto wks = doc.workbook().worksheet("Sheet1");
        wks.cell("A1").value() = "This file was encrypted by OpenXLSX C++ Engine";
        wks.cell("B1").value() = 42;
        wks.cell("C1").formula() = "B1*2";
        
        // Save the document with a password
        REQUIRE_NOTHROW(doc.saveAs("Tests/Fixtures/OpenXLSX_Encrypted_Cycle_Test.xlsx", std::string("CyclePass123!"), XLForceOverwrite));
        
        // Verify it can be opened back
        XLDocument doc2;
        REQUIRE_NOTHROW(doc2.open("Tests/Fixtures/OpenXLSX_Encrypted_Cycle_Test.xlsx", "CyclePass123!"));
        
        // Verify the contents
        auto wks2 = doc2.workbook().worksheet("Sheet1");
        REQUIRE(wks2.cell("A1").value().getString() == "This file was encrypted by OpenXLSX C++ Engine");
        REQUIRE(wks2.cell("B1").value().get<int>() == 42);
        REQUIRE(wks2.cell("C1").formula().get() == "B1*2");
    }

    SECTION("Invalid Password Handling") {
        XLDocument doc;
        doc.create("Tests/Fixtures/OpenXLSX_Wrong_Pass_Test.xlsx", XLForceOverwrite);
        doc.workbook().worksheet("Sheet1").cell("A1").value() = "Secret Data";
        doc.saveAs("Tests/Fixtures/OpenXLSX_Wrong_Pass_Test.xlsx", std::string("CorrectPass"), XLForceOverwrite);

        XLDocument doc2;
        // Opening with the wrong password should throw an XLInternalError (due to HMAC/Verifier mismatch or ZIP signature mismatch)
        REQUIRE_THROWS_AS(doc2.open("Tests/Fixtures/OpenXLSX_Wrong_Pass_Test.xlsx", "WrongPass"), XLInternalError);
    }
}
