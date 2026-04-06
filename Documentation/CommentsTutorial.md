# Quick Start: Cell Comments & Threads {#comments_tutorial}

[TOC]

## Introduction

OpenXLSX supports two different types of cell annotations provided by Microsoft Excel:
1. **Traditional Comments (Notes):** The classic floating yellow boxes with a red triangle indicator in the cell corner. These are typically used for static annotations or instructions.
2. **Modern Threaded Comments:** The modern conversational UI (introduced in Office 365) with a purple indicator. These allow multiple users to have a back-and-forth discussion directly tied to a cell.

This tutorial covers:
- Adding, sizing, and attributing traditional comments.
- Initiating modern threaded conversations.
- Adding replies to existing conversation threads.
- Removing comments programmatically.

## 1. Traditional Comments (Notes)

To add traditional comments, you interact with the `comments()` collection on the worksheet. First, you register an author name, which returns an author ID. You then use `set()` to attach a comment to a specific cell reference.

```cpp
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("CommentsDemo.xlsx", XLForceOverwrite);
    
    auto wks = doc.workbook().worksheet("Sheet1");
    wks.setName("Traditional Comments");

    // 1. Register authors (returns a 0-based author ID)
    uint16_t authorAdmin = wks.comments().addAuthor("System Admin");
    uint16_t authorAudit = wks.comments().addAuthor("Financial Auditor");

    // 2. Add a standard-sized comment to a cell
    wks.cell("A2").value() = "Default Size Comment";
    wks.comments().set("A2", "This is a standard sized note.", authorAdmin);

    // 3. Add a large comment 
    // The set() method accepts optional width (in columns) and height (in rows)
    wks.cell("B2").value() = "Large Comment";
    wks.comments().set("B2", "This comment is much larger to accommodate\nmultiple lines\nof text.", authorAdmin, 6, 10);
    
    // 4. Add a comment by a different author
    wks.cell("C2").value() = "Auditor Comment";
    wks.comments().set("C2", "Please review these numbers.", authorAudit);
```

## 2. Modern Threaded Comments

For modern conversational comments, OpenXLSX provides highly streamlined, integrated methods directly on the `XLWorksheet` object: `addThreadedComment()` and `addThreadedReply()`. 

These methods automatically manage the underlying XML structures and "Person" identities required by modern Excel versions.

```cpp
    doc.workbook().addWorksheet("Threaded Comments");
    auto wksThreaded = doc.workbook().worksheet("Threaded Comments");

    wksThreaded.cell("A3").value() = "Q3 Projections";
    
    // 1. Start a new conversation thread on a cell
    // Signature: addThreadedComment(CellReference, Text, AuthorName)
    // Returns an XLThreadedComment object which contains the unique thread ID.
    auto thread1 = wksThreaded.addThreadedComment("A3", "Can we update the Q3 projections?", "Alice Smith");
    
    // 2. Add replies to the conversation thread
    // Signature: addThreadedReply(ParentThreadID, Text, AuthorName)
    wksThreaded.addThreadedReply(thread1.id(), "Yes, I will have them ready by tomorrow.", "Bob Johnson");
    wksThreaded.addThreadedReply(thread1.id(), "Great, thanks Bob!", "Alice Smith");


    wksThreaded.cell("C3").value() = "Revenue Figures";
    
    // 3. Start a separate conversation thread on a different cell
    auto thread2 = wksThreaded.addThreadedComment("C3", "Is this figure correct?", "Bob Johnson");
    wksThreaded.addThreadedReply(thread2.id(), "I believe so, let me double check the raw data.", "Alice Smith");
```

## 3. Removing Comments

You can programmatically remove comments and entire threaded conversations from cells using the corresponding deletion methods.

```cpp
    // Remove a legacy traditional comment
    wks.deleteComment("B2");

    // Remove a modern threaded comment AND all of its replies
    wksThreaded.deleteThreadedComment("C3");
```

## Summary

Save and close your document. 

```cpp
    doc.save();
    doc.close();
    
    return 0;
}
```

The generated file will natively display the red/purple indicators and provide the rich interaction features users expect within Microsoft Excel.
