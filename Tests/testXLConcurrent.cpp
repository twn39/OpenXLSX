#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <thread>
#include <vector>
#include <atomic>

using namespace OpenXLSX;

TEST_CASE("XLConcurrent Tests", "[XLConcurrent]")
{
    SECTION("Concurrent writes to different worksheets")
    {
        const std::string filename = "./testXLConcurrent.xlsx";
        constexpr int     numSheets = 4;
        constexpr int     numRows   = 200;

        // Setup: create document with multiple worksheets
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        for (int i = 2; i <= numSheets; ++i) {
            doc.workbook().addWorksheet("Sheet" + std::to_string(i));
        }

        // Get worksheet references before spawning threads
        std::vector<XLWorksheet> worksheets;
        for (int i = 1; i <= numSheets; ++i) {
            worksheets.push_back(doc.workbook().worksheet("Sheet" + std::to_string(i)));
        }

        // Concurrent writes: each thread writes to its own worksheet
        std::vector<std::thread> threads;
        std::atomic<int>         errors{0};

        for (int t = 0; t < numSheets; ++t) {
            threads.emplace_back([&, t]() {
                try {
                    auto& wks = worksheets[static_cast<size_t>(t)];
                    for (int row = 1; row <= numRows; ++row) {
                        wks.cell(static_cast<uint32_t>(row), 1).value() = row;                                              // integer
                        wks.cell(static_cast<uint32_t>(row), 2).value() = "Sheet" + std::to_string(t + 1) + "_Row" + std::to_string(row);  // string (uses shared strings)
                        wks.cell(static_cast<uint32_t>(row), 3).value() = row * 1.5;                                        // double
                    }
                }
                catch (...) {
                    ++errors;
                }
            });
        }

        for (auto& th : threads) { th.join(); }
        REQUIRE(errors == 0);

        // Save should work without deadlock
        doc.save();

        // Verify data integrity
        doc.close();
        doc.open(filename);

        for (int t = 0; t < numSheets; ++t) {
            auto wks = doc.workbook().worksheet("Sheet" + std::to_string(t + 1));
            REQUIRE(wks.cell(1, 1).value().get<int>() == 1);
            REQUIRE(wks.cell(numRows, 1).value().get<int>() == numRows);
            REQUIRE(wks.cell(1, 2).value().get<std::string>() == "Sheet" + std::to_string(t + 1) + "_Row1");
            REQUIRE(wks.cell(1, 3).value().get<double>() == Catch::Approx(1.5));
        }

        doc.close();
    }

    SECTION("Concurrent reads after writes")
    {
        const std::string filename = "./testXLConcurrentRead.xlsx";
        constexpr int     numRows  = 100;

        // Setup: create and populate
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            doc.workbook().addWorksheet("Sheet2");

            auto wks1 = doc.workbook().worksheet("Sheet1");
            auto wks2 = doc.workbook().worksheet("Sheet2");
            for (int row = 1; row <= numRows; ++row) {
                wks1.cell(static_cast<uint32_t>(row), 1).value() = row;
                wks2.cell(static_cast<uint32_t>(row), 1).value() = row * 10;
            }
            doc.save();
            doc.close();
        }

        // Concurrent reads
        XLDocument       doc;
        doc.open(filename);
        std::atomic<int> errors{0};

        std::thread reader1([&]() {
            try {
                auto wks = doc.workbook().worksheet("Sheet1");
                for (int row = 1; row <= numRows; ++row) {
                    int val = wks.cell(static_cast<uint32_t>(row), 1).value().get<int>();
                    if (val != row) ++errors;
                }
            }
            catch (...) { ++errors; }
        });

        std::thread reader2([&]() {
            try {
                auto wks = doc.workbook().worksheet("Sheet2");
                for (int row = 1; row <= numRows; ++row) {
                    int val = wks.cell(static_cast<uint32_t>(row), 1).value().get<int>();
                    if (val != row * 10) ++errors;
                }
            }
            catch (...) { ++errors; }
        });

        reader1.join();
        reader2.join();
        REQUIRE(errors == 0);
        doc.close();
    }

    SECTION("Shared string table thread safety")
    {
        const std::string filename = "./testXLConcurrentSST.xlsx";
        constexpr int     numThreads = 4;
        constexpr int     numStrings = 100;

        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        for (int i = 2; i <= numThreads; ++i) {
            doc.workbook().addWorksheet("Sheet" + std::to_string(i));
        }

        std::vector<XLWorksheet> worksheets;
        for (int i = 1; i <= numThreads; ++i) {
            worksheets.push_back(doc.workbook().worksheet("Sheet" + std::to_string(i)));
        }

        // All threads write the SAME strings to force shared string contention
        std::vector<std::thread> threads;
        std::atomic<int>         errors{0};

        for (int t = 0; t < numThreads; ++t) {
            threads.emplace_back([&, t]() {
                try {
                    auto& wks = worksheets[static_cast<size_t>(t)];
                    for (int i = 1; i <= numStrings; ++i) {
                        // All threads write the same string values → contention on shared string table
                        wks.cell(static_cast<uint32_t>(i), 1).value() = "CommonString_" + std::to_string(i);
                    }
                }
                catch (...) { ++errors; }
            });
        }

        for (auto& th : threads) { th.join(); }
        REQUIRE(errors == 0);

        doc.save();

        // Verify: all sheets should have the same data
        doc.close();
        doc.open(filename);

        for (int t = 0; t < numThreads; ++t) {
            auto wks = doc.workbook().worksheet("Sheet" + std::to_string(t + 1));
            for (int i = 1; i <= numStrings; ++i) {
                REQUIRE(wks.cell(static_cast<uint32_t>(i), 1).value().get<std::string>() == "CommonString_" + std::to_string(i));
            }
        }
        doc.close();
    }
}
