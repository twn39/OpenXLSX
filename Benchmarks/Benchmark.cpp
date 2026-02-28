#include <OpenXLSX.hpp>
#include <catch2/catch_test_macros.hpp>
#include <catch2/benchmark/catch_benchmark.hpp>
#include <deque>
#include <fstream>
#include <numeric>
#include <vector>

using namespace OpenXLSX;

// Using a slightly smaller row count for regular benchmarking to avoid excessive runtimes, 
// but enough to show performance characteristics.
constexpr uint64_t rowCount = 100000; 
constexpr uint8_t  colCount = 8;

TEST_CASE("OpenXLSX Benchmarks", "[.benchmark]")
{
    SECTION("Write Operations")
    {
        BENCHMARK("Write Strings")
        {
            XLDocument doc;
            doc.create("./benchmark_strings.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            std::deque<XLCellValue> values(colCount, "OpenXLSX");

            for (auto& row : wks.rows(rowCount)) row.values() = values;

            doc.save();
            doc.close();
            return rowCount * colCount;
        };

        BENCHMARK("Write Integers")
        {
            XLDocument doc;
            doc.create("./benchmark_integers.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            std::vector<XLCellValue> values(colCount, 42);

            for (auto& row : wks.rows(rowCount)) row.values() = values;

            doc.save();
            doc.close();
            return rowCount * colCount;
        };

        BENCHMARK("Write Floats")
        {
            XLDocument doc;
            doc.create("./benchmark_floats.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            std::vector<XLCellValue> values(colCount, 3.14);

            for (auto& row : wks.rows(rowCount)) row.values() = values;

            doc.save();
            doc.close();
            return rowCount * colCount;
        };

        BENCHMARK("Write Bools")
        {
            XLDocument doc;
            doc.create("./benchmark_bools.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            std::vector<XLCellValue> values(colCount, true);

            for (auto& row : wks.rows(rowCount)) row.values() = values;

            doc.save();
            doc.close();
            return rowCount * colCount;
        };
    }

    SECTION("Read Operations")
    {
        // Prerequisites for read benchmarks (ensure files exist)
        {
            XLDocument doc;
            if (!std::ifstream("./benchmark_strings.xlsx")) {
                doc.create("./benchmark_strings.xlsx", XLForceOverwrite);
                auto wks = doc.workbook().worksheet("Sheet1");
                std::deque<XLCellValue> values(colCount, "OpenXLSX");
                for (auto& row : wks.rows(rowCount)) row.values() = values;
                doc.save();
            }
            if (!std::ifstream("./benchmark_integers.xlsx")) {
                doc.create("./benchmark_integers.xlsx", XLForceOverwrite);
                auto wks = doc.workbook().worksheet("Sheet1");
                std::vector<XLCellValue> values(colCount, 42);
                for (auto& row : wks.rows(rowCount)) row.values() = values;
                doc.save();
            }
        }

        BENCHMARK("Read Strings")
        {
            XLDocument doc;
            doc.open("./benchmark_strings.xlsx");
            auto     wks    = doc.workbook().worksheet("Sheet1");
            uint64_t result = 0;

            for (auto& row : wks.rows()) {
                std::vector<std::string> values = std::vector<std::string>(row.values());
                result += std::count_if(values.begin(), values.end(), [](const std::string& v) {
                    return !v.empty();
                });
            }
            doc.close();
            return result;
        };

        BENCHMARK("Read Integers")
        {
            XLDocument doc;
            doc.open("./benchmark_integers.xlsx");
            auto     wks    = doc.workbook().worksheet("Sheet1");
            uint64_t result = 0;

            for (auto& row : wks.rows()) {
                std::vector<XLCellValue> values = row.values();
                result += std::accumulate(values.begin(),
                                          values.end(),
                                          0ULL,
                                          [](uint64_t a, const XLCellValue& b) { return a + b.get<uint64_t>(); });
            }
            doc.close();
            return result;
        };
    }
}
