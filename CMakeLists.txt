cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX VERSION 0.4.1 LANGUAGES CXX)

# Set C++ standard
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# Build options
option(OPENXLSX_BUILD_TESTS "Build and run library tests" ON)
option(OPENXLSX_BUILD_BENCHMARKS "Build benchmark programs" OFF)
option(OPENXLSX_CREATE_DOCS "Build library documentation" OFF)
option(OPENXLSX_ENABLE_NOWIDE "Enable Nowide for UTF-8 support on Windows" ON)
option(OPENXLSX_ENABLE_LTO "Enable Link-Time Optimization (LTO/IPO) if supported" ON)

set(OPENXLSX_LIBRARY_TYPE "STATIC" CACHE STRING "Library type (STATIC or SHARED)")

#=======================================================================================================================
# External Libraries
#=======================================================================================================================
include(FetchContent)

# Set up fast_float
FetchContent_Declare(fast_float GIT_REPOSITORY https://github.com/fastfloat/fast_float.git GIT_TAG v8.2.3)
FetchContent_MakeAvailable(fast_float)

# Set up fmt
FetchContent_Declare(fmt GIT_REPOSITORY https://github.com/fmtlib/fmt.git GIT_TAG 12.1.0)
FetchContent_MakeAvailable(fmt)

# Set up GSL
FetchContent_Declare(GSL GIT_REPOSITORY https://github.com/microsoft/GSL.git GIT_TAG v4.2.1)
FetchContent_MakeAvailable(GSL)

# Disable testing for third-party dependencies
set(BUILD_TESTING_OLD ${BUILD_TESTING})
set(BUILD_TESTING OFF CACHE BOOL "" FORCE)

# Set up zlib-ng
FetchContent_Declare(zlib_ng GIT_REPOSITORY https://github.com/zlib-ng/zlib-ng.git GIT_TAG 2.2.4)
set(ZLIB_COMPAT ON CACHE BOOL "Enable zlib compatibility mode")
set(ZLIB_ENABLE_SIMD ON CACHE BOOL "Enable SIMD optimizations")
set(ZLIBNG_ENABLE_TESTS OFF CACHE BOOL "Disable zlib-ng tests")
set(ZLIB_ENABLE_TESTS OFF CACHE BOOL "Disable zlib compat tests")
set(WITH_GTEST OFF CACHE BOOL "Disable zlib-ng gtest")

FetchContent_GetProperties(zlib_ng)
if(NOT zlib_ng_POPULATED)
    FetchContent_Populate(zlib_ng)
    add_subdirectory(${zlib_ng_SOURCE_DIR} ${zlib_ng_BINARY_DIR} EXCLUDE_FROM_ALL)
endif()

# Set up minizip-ng
FetchContent_Declare(minizip_ng GIT_REPOSITORY https://github.com/zlib-ng/minizip-ng.git GIT_TAG 4.0.8)
set(MZ_COMPAT ON CACHE BOOL "Enable zlib compatibility mode")
set(MZ_ZLIB ON CACHE BOOL "Enable ZLIB support")
set(MZ_BZIP2 OFF CACHE BOOL "Disable BZIP2 support")
set(MZ_LZMA OFF CACHE BOOL "Disable LZMA support")
set(MZ_PKCRYPT OFF CACHE BOOL "Disable PKWARE encryption")
set(MZ_WZAES OFF CACHE BOOL "Disable WinZAES encryption")
set(MZ_LIBCOMP OFF CACHE BOOL "Disable Apple Compression Framework")
set(MZ_ZSTD OFF CACHE BOOL "Disable ZSTD support")
set(MZ_OPENSSL OFF CACHE BOOL "Disable OpenSSL support")
set(MZ_ICONV OFF CACHE BOOL "Disable ICONV support")
set(MZ_FETCH_LIBS OFF CACHE BOOL "Disable fetching third-party libs")
set(MZ_BUILD_TESTS OFF CACHE BOOL "Disable minizip-ng tests")
set(MZ_BUILD_UNIT_TESTS OFF CACHE BOOL "Disable minizip-ng unit tests")

FetchContent_GetProperties(minizip_ng)
if(NOT minizip_ng_POPULATED)
    FetchContent_Populate(minizip_ng)
    add_subdirectory(${minizip_ng_SOURCE_DIR} ${minizip_ng_BINARY_DIR} EXCLUDE_FROM_ALL)
endif()

# Restore testing flag
set(BUILD_TESTING ${BUILD_TESTING_OLD} CACHE BOOL "" FORCE)

# Set up Catch2 (if tests or benchmarks are enabled)
if(OPENXLSX_BUILD_TESTS OR OPENXLSX_BUILD_BENCHMARKS)
    FetchContent_Declare(Catch2 GIT_REPOSITORY https://github.com/catchorg/Catch2.git GIT_TAG v3.13.0)
    FetchContent_MakeAvailable(Catch2)
endif()

# Third-party subdirectories
set(PUGIXML_XPATH OFF CACHE BOOL "" FORCE)
add_subdirectory(third_party/pugixml EXCLUDE_FROM_ALL)

set(NOWIDE_INSTALL OFF CACHE BOOL "" FORCE)
set(NOWIDE_TESTING OFF CACHE BOOL "" FORCE)
add_subdirectory(third_party/nowide EXCLUDE_FROM_ALL)

if(WIN32)
    set(OPENXLSX_ENABLE_NOWIDE ON)
endif()

# Sources
file(GLOB OPENXLSX_SOURCES OpenXLSX/sources/*.cpp)

# Library Target
add_library(OpenXLSX ${OPENXLSX_LIBRARY_TYPE} ${OPENXLSX_SOURCES})
add_library(OpenXLSX::OpenXLSX ALIAS OpenXLSX)

target_include_directories(OpenXLSX
    PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/OpenXLSX>
        $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/OpenXLSX/headers>
        $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>
    SYSTEM PUBLIC
        $<BUILD_INTERFACE:${fast_float_SOURCE_DIR}/include>
        $<BUILD_INTERFACE:${fmt_SOURCE_DIR}/include>
        $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/third_party/pugixml/src>
        $<BUILD_INTERFACE:${minizip_ng_SOURCE_DIR}>
)

# On Windows, minizip might be named minizip_static or minizip
if(TARGET minizip)
    set(MINIZIP_TARGET minizip)
elseif(TARGET minizip_static)
    set(MINIZIP_TARGET minizip_static)
else()
    set(MINIZIP_TARGET minizip)
endif()

target_link_libraries(OpenXLSX PUBLIC pugixml Microsoft.GSL::GSL ${MINIZIP_TARGET})

if(WIN32)
    # minizip-ng on Windows needs these for random number generation and OS operations
    target_link_libraries(OpenXLSX PRIVATE Bcrypt User32 Advapi32)
    target_compile_definitions(OpenXLSX PUBLIC _CRT_SECURE_NO_WARNINGS)
endif()

target_compile_definitions(OpenXLSX PUBLIC FMT_HEADER_ONLY)

if (OPENXLSX_ENABLE_NOWIDE)
    target_link_libraries(OpenXLSX PRIVATE nowide::nowide)
    target_compile_definitions(OpenXLSX PRIVATE ENABLE_NOWIDE)
endif ()

if ("${OPENXLSX_LIBRARY_TYPE}" STREQUAL "STATIC")
    target_compile_definitions(OpenXLSX PUBLIC OPENXLSX_STATIC_DEFINE)
endif ()

# Enable Link-Time Optimization (LTO/IPO) if supported
if(OPENXLSX_ENABLE_LTO)
    include(CheckIPOSupported)
    check_ipo_supported(RESULT result OUTPUT output)
    if(result)
        set_property(TARGET OpenXLSX PROPERTY INTERPROCEDURAL_OPTIMIZATION TRUE)
    endif()
endif()

# Export Header (Required by headers)
include(GenerateExportHeader)
generate_export_header(OpenXLSX
    BASE_NAME openxlsx
    EXPORT_FILE_NAME OpenXLSX-Exports.hpp
    EXPORT_MACRO_NAME OPENXLSX_EXPORT
)

# Basic Compiler Flags & Release Optimizations
if(MSVC)
    target_compile_options(OpenXLSX PUBLIC /W4 /permissive- /utf-8 $<$<CONFIG:Release>:/O2 /Ot /Gy /Zc:inline>)
    target_compile_definitions(OpenXLSX PUBLIC _CRT_SECURE_NO_WARNINGS)
else()
    target_compile_options(OpenXLSX PRIVATE -Wall -Wextra $<$<CONFIG:Release>:-O3 -ffunction-sections -fdata-sections>)
    if(APPLE)
        target_link_options(OpenXLSX PRIVATE $<$<CONFIG:Release>:-Wl,-dead_strip>)
    else()
        target_link_options(OpenXLSX PRIVATE $<$<CONFIG:Release>:-Wl,--gc-sections>)
    endif()
endif()

#=======================================================================================================================
# TESTS (Merged configuration)
#=======================================================================================================================
if(OPENXLSX_BUILD_TESTS)
    file(GLOB TEST_SOURCES Tests/testXL*.cpp Tests/main.cpp)
    add_executable(OpenXLSXTests ${TEST_SOURCES})
    target_link_libraries(OpenXLSXTests PRIVATE OpenXLSX::OpenXLSX Catch2::Catch2WithMain)

    # Copy test data to build directory
    add_custom_command(TARGET OpenXLSXTests POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:OpenXLSXTests>/Tests
        COMMAND ${CMAKE_COMMAND} -E make_directory ${CMAKE_BINARY_DIR}/Tests
        COMMAND ${CMAKE_COMMAND} -E copy_if_different 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/test.png 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/test.jpg
            ${CMAKE_CURRENT_LIST_DIR}/Tests/shared_formula_test.xlsx
            $<TARGET_FILE_DIR:OpenXLSXTests>/Tests/
        COMMAND ${CMAKE_COMMAND} -E copy_if_different 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/test.png 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/test.jpg
            ${CMAKE_CURRENT_LIST_DIR}/Tests/shared_formula_test.xlsx
            ${CMAKE_BINARY_DIR}/Tests/
    )
endif()

#=======================================================================================================================
# BENCHMARKS (Merged configuration using Catch2)
#=======================================================================================================================
if(OPENXLSX_BUILD_BENCHMARKS)
    add_executable(OpenXLSXBenchmark Benchmarks/Benchmark.cpp)
    target_link_libraries(OpenXLSXBenchmark PRIVATE OpenXLSX::OpenXLSX Catch2::Catch2WithMain)
endif()

#=======================================================================================================================
# DOCUMENTATION (Subdirectory configuration)
#=======================================================================================================================
if(OPENXLSX_CREATE_DOCS)
    add_subdirectory(Documentation)
endif()
