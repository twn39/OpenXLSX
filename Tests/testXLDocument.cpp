//
// Created by Troldal on 2019-01-12.
//

#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <fstream>

using namespace OpenXLSX;

/**
 * @brief The purpose of this test case is to test the creation of XLDocument objects. Each section section
 * tests document creation using a different method. In addition, saving, closing and copying is tested.
 */
TEST_CASE("XLDocument Tests", "[XLDocument]")
{
    std::string file    = "./testXLDocument.xlsx";
    std::string newfile = "./TestDocumentCreationNew.xlsx";

    /**
     * @test
     *
     * @details
     */
    SECTION("Create empty XLDocument, using default constructor")
    {
        XLDocument doc;
        REQUIRE_FALSE(doc);
    }

    SECTION("Create new using create()")
    {
        XLDocument doc;
        doc.create(file, XLForceOverwrite);
        std::ifstream f(file);
        REQUIRE(f.good());
        REQUIRE(doc.name() == "testXLDocument.xlsx");
        doc.close();
    }

    SECTION("Open existing using Constructor")
    {
        {
            XLDocument doc1;
            doc1.create(file, XLForceOverwrite);
            doc1.close();
        }
        XLDocument doc2(file);
        REQUIRE(doc2.name() == "testXLDocument.xlsx");
        doc2.close();
    }

    SECTION("Open existing using open()")
    {
        {
            XLDocument doc1;
            doc1.create(file, XLForceOverwrite);
            doc1.close();
        }
        XLDocument doc2;
        doc2.open(file);
        REQUIRE(doc2.name() == "testXLDocument.xlsx");
        doc2.close();
    }

    SECTION("Save document using saveAs()")
    {
        {
            XLDocument doc1;
            doc1.create(file, XLForceOverwrite);
            doc1.close();
        }
        XLDocument doc2(file);
        doc2.saveAs(newfile, XLForceOverwrite);
        std::ifstream n(newfile);
        REQUIRE(n.good());
        REQUIRE(doc2.name() == "TestDocumentCreationNew.xlsx");
        doc2.close();
    }

    SECTION("Refuse to overwrite existing file")
    {
        {
            XLDocument doc1;
            doc1.create(file, XLForceOverwrite);
            doc1.close();
        }

        XLDocument doc2;
        REQUIRE_THROWS_AS(doc2.create(file, XLDoNotOverwrite), XLException);
    }

    SECTION("UTF-8 file paths and names")
    {
        std::string utf8File = "./测试_test_中文.xlsx";
        std::remove(utf8File.c_str());

        {
            XLDocument doc;
            doc.create(utf8File, XLForceOverwrite);
            doc.workbook().worksheet("Sheet1").cell("A1").value() = "UTF8 Path Test";
            doc.save();
            doc.close();
        }

        // Verify it can be opened
        {
            XLDocument doc;
            doc.open(utf8File);
            REQUIRE(doc.workbook().worksheet("Sheet1").cell("A1").value().get<std::string>() == "UTF8 Path Test");
            doc.close();
        }

        // Test saveAs with UTF-8
        {
            std::string utf8File2 = "./测试_saveAs_中文.xlsx";
            std::remove(utf8File2.c_str());
            XLDocument doc;
            doc.open(utf8File);
            doc.saveAs(utf8File2, XLForceOverwrite);
            doc.close();

            XLDocument doc2;
            doc2.open(utf8File2);
            REQUIRE(doc2.isOpen());
            doc2.close();
            std::remove(utf8File2.c_str());
        }

        std::remove(utf8File.c_str());
    }
}
