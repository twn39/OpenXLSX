//
// Created by Troldal on 2019-01-12.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
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
    }
    