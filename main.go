package main

import (
	"fmt"
	"github.com/gin-gonic/gin"
	"github.com/tealeg/xlsx"
)

func main() {
	r := gin.Default()

	r.GET("/generate-xlsx", func(c *gin.Context) {
		generateAndSaveXsl("doc.xlsx")
		c.File("doc.xlsx")
	})

	r.Run() // listen and serve on 0.0.0.0:8080
}

func generateAndSaveXsl(fileName string) {
	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Println(err.Error())
	}

	style := &xlsx.Style{
		Alignment: xlsx.Alignment{
			Horizontal: "center",
			Vertical:   "bottom",
		},
		Border: xlsx.Border{
			Left:   "thin",
			Right:  "thin",
			Top:    "thin",
			Bottom: "thin",
		},
		Fill: xlsx.Fill{
			PatternType: xlsx.Solid_Cell_Fill,
			FgColor:     xlsx.RGB_Light_Green,
			BgColor:     xlsx.RGB_Light_Green,
		},
		Font: xlsx.Font{Size: 12, Name: xlsx.Helvetica},
	}

	style.Font.Bold = true

	for i := 0; i < 2000; i++ {
		row = sheet.AddRow()
		cell = row.AddCell()

		for j := 0; j < 10; j++ {
			cell = row.GetCell(j)
			row.GetCell(j).Value = "111"
			cell.SetStyle(style)
		}
	}

	err = file.Save(fileName)
	if err != nil {
		fmt.Println(err.Error())
	}
}
