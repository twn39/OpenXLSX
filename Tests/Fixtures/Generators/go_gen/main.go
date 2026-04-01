package main

import (
	"fmt"
	"time"
	"github.com/xuri/excelize/v2"
)

func main() {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	f.SetSheetName("Sheet1", "TestSheet")
	
	// 1. Basic Values
	f.SetCellValue("TestSheet", "A1", "Hello from excelize")
	f.SetCellValue("TestSheet", "A2", 42)
	f.SetCellValue("TestSheet", "A3", 3.1415)

	// 2. Date and Time
	f.SetCellValue("TestSheet", "A4", time.Date(2024, 5, 28, 14, 30, 0, 0, time.UTC))

	// 3. Formulas
	f.SetCellFormula("TestSheet", "A5", "=SUM(A2:A3)")

	// 4. Styling (Font, Fill, Border, Alignment)
	f.SetCellValue("TestSheet", "B1", "Styled")
	style, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Italic: true, Color: "FF0000", Family: "Comic Sans MS"},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"00FF00"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	if err == nil {
		f.SetCellStyle("TestSheet", "B1", "B1", style)
	}

	// 5. Merged Cells
	f.MergeCell("TestSheet", "C1", "E2")
	f.SetCellValue("TestSheet", "C1", "Merged Area")

	// 6. Multiple Sheets
	f.NewSheet("SecondSheet")
	f.SetCellValue("SecondSheet", "A1", "Data in second sheet")

	if err := f.SaveAs("../../excelize_generated.xlsx"); err != nil {
		fmt.Println(err)
	}
}

