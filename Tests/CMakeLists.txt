#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX.Tests)

#=======================================================================================================================
# Set up Catch2 library (v3)
#=======================================================================================================================
include(FetchContent)
FetchContent_Declare(
  Catch2
  GIT_REPOSITORY https://github.com/catchorg/Catch2.git
  GIT_TAG        v3.13.0
)
FetchContent_MakeAvailable(Catch2)

#=======================================================================================================================
# Define TEST target
#=======================================================================================================================
add_executable(OpenXLSXTests EXCLUDE_FROM_ALL "")
target_sources(OpenXLSXTests
        PRIVATE
        main.cpp
        testXLDocument.cpp
        testXLCell.cpp
        testXLCellRange.cpp
        testXLCellReference.cpp
        testXLCellValue.cpp
        testXLCellValueProxy.cpp
        testXLColor.cpp
        testXLColumn.cpp
        testXLComments.cpp
        testXLDateTime.cpp
        testXLFormula.cpp
        testXLHyperlink.cpp
        testXLImage.cpp
        testXLIntegration.cpp
        testXLMergeCells.cpp
        testXLRow.cpp
        testXLSheet.cpp
        testXLSheetExtended.cpp
        testXLSharedStrings.cpp
        testXLStyles.cpp
        testXLStylesIntegration.cpp
        testXLDrawingTables.cpp
        testXLTables.cpp
        testXLWorkbookProperties.cpp
        )

target_link_libraries(OpenXLSXTests
        PRIVATE
        OpenXLSX::OpenXLSX
        Catch2::Catch2WithMain
        )