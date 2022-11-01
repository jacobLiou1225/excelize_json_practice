package main

import (
	"fmt"

	_ "image/jpeg"
	_ "image/png"

	"github.com/xuri/excelize/v2"
)

func main() {
	f := excelize.NewFile()
	sheetName := "成績單"
	f.SetSheetName("Sheet1", sheetName)
	data := [][]interface{}{
		{"考試成績統計表"},
		{"考試名稱:期中考", nil, nil, nil, "基礎科目", nil, nil, "術科科目"},
		{"順序", "學號", "姓名", "性別", "math", "english", "chinese", "生物", "物理", "化學"},
		{1, 1001, "大熊", "男", 95, 63, 25, 78, 85, 34},
		{2, 1002, "小熊", "男", 95, 10, 25, 67, 14, 34},
		{3, 1003, "中熊", "男", 95, 63, 25, 52, 78, 87},
		{4, 1004, "阿熊", "男", 90, 63, 80, 78, 20, 57},
		{5, 1005, "棒熊", "?", 95, 63, 45, 42, 85, 68},
	}
	for i, row := range data {
		startCell, err := excelize.JoinCellName("A", i+1)
		if err != nil {
			fmt.Print(err)
			return
		}
		if err := f.SetSheetRow(sheetName, startCell, &row); err != nil {
			fmt.Print(err)
			return
		}

	}
	//加總公式設定
	formularType, ref := excelize.STCellFormulaTypeShared, "K4:K8"
	if err := f.SetCellFormula(sheetName, "K4", "=SUM(E4:J4)",
		excelize.FormulaOpts{Ref: &ref, Type: &formularType}); err != nil {
		fmt.Println(err)
		return
	}
	//合併欄位
	mergeCellRange := [][]string{{"A1", "K1"}, {"A2", "D2"}, {"E2", "G2"}, {"H2", "J2"}}
	for _, rangs := range mergeCellRange {
		if err := f.MergeCell(sheetName, rangs[0], rangs[1]); err != nil {
			fmt.Println(err)
		}
	}
	//設定title顏色跟字形置中
	style1, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#DFEBF6"}, Pattern: 1},
	})
	if err != nil {
		fmt.Println(err)
		return
	}

	if f.SetCellStyle(sheetName, "A1", "A1", style1); err != nil {
		fmt.Println(err)
		return
	}
	//標題置中
	style2, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center"},
	})
	if err != nil {
		fmt.Println(err)
		return
	}
	for _, cell := range []string{"A2", "E2", "H2"} {
		if f.SetCellStyle(sheetName, cell, cell, style2); err != nil {
			fmt.Println(err)
			return
		}
	}
	//行寬
	if err := f.SetColWidth(sheetName, "D", "k", 7); err != nil {
		fmt.Println(err)
	}

	//建立表格 上顏色
	if err := f.AddTable(sheetName, "A3", "J8", `{
		
		"table_name":"table",
		"table_style":"TableStyleLight2"
	}`); err != nil {
		fmt.Println(err)
		return
	}

	//家圖片
	if err := f.AddPicture(sheetName, "G9", "cat.png", `{
		"x_offset":15,
		"y_offset":15,
		"x_scale":0.2,
		"y_scale":0.2
	}`); err != nil {
		fmt.Println(err)
		return
	}
	//將表格線隱藏
	if err := f.SetSheetViewOptions(sheetName, 0,
		excelize.ShowGridLines(false)); err != nil {
		fmt.Println(err)
		return
	}
	//將數據鎖住
	if err := f.SetPanes(sheetName, `{
		"freeze":true,
		"split":false,
		"x_split":0,
		"y_split":3,
		"top_left_cell":"A4",
		"active_pane": "bottomLeft"
		}`); err != nil {
		fmt.Println(err)
		return
	}

	//加圖表
	if err := f.AddChartSheet("設計圖", `{
		"type" : "col",
		"series" : [
			{
				"name" : "成績單!$A$2",
				"categories": "成績單!$C$4:$C$8",
				"values": "成績單!$K$4:$K$8"
			}
		],
		"plotarea" :{
			"show_val":true
		},
		
		"legend" : {
			"none" : true
		},
		"title" : {
			"name" : "成績單"
		}
	}`); err != nil {
		fmt.Println(err)
		return
	}

	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Print(err)
		return
	}
}
