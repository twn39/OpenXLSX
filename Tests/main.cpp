//
// Created by Troldal on 2019-01-10.
//

#ifndef OPENXLSX_TESTMAIN_H
#define OPENXLSX_TESTMAIN_H

#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <cstdio>

using namespace OpenXLSX;

void PrepareDocument(std::string name)
{
    XLDocument doc;
    std::remove(name.c_str());
    doc.create(name, XLForceOverwrite);
    doc.save();
    doc.close();
}

#endif    // OPENXLSX_TESTMAIN_H
