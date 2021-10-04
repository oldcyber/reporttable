package lib

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// TODO: Обрезать имя любого столбца
func ConvStrInt(str []string) []int {
	var curRes []string
	var curIntRes = []int{}
	for a := 0; a < len(str); a++ {
		curRes = append(curRes, strings.TrimPrefix(str[a], "B"))
	}

	for _, a := range curRes {
		b, err := strconv.Atoi(a)
		if err != nil {
			panic(err)
		}
		curIntRes = append(curIntRes, b)
	}
	return curIntRes
}

func SetRowsColor(firstSheet string, rows int, cols int, level int, f *excelize.File) {
	
	// Определяем стили
	styleGrey, err := f.NewStyle(`{"fill":{"type":"pattern","color":["#BFBFBF"],"pattern":1}}`)
	if err != nil {
		fmt.Println(err)
	}
	styleBlue, err := f.NewStyle(`{"fill":{"type":"pattern","color":["#95B3D7"],"pattern":1}}`)
	if err != nil {
		fmt.Println(err)
	}

	var firstCell string = "A" + strconv.Itoa(rows)
	c, err := excelize.ColumnNumberToName(cols)
	if err != nil {
		fmt.Println(err)
	}
	var lastCell string = c + strconv.Itoa(rows)
	if level == 1 {

		//красим в серый цвет
		err = f.SetCellStyle(firstSheet, firstCell, lastCell, styleGrey)
		if err != nil {
			fmt.Println(err)
		}
	} else if level == 2 {
		//красим в голубой цвет
		err = f.SetCellStyle(firstSheet, firstCell, lastCell, styleBlue)
		if err != nil {
			fmt.Println(err)
		}
	}
	//else {
	//}

	// Область печати
	// var areas string = firstSheet + "!$A$1:$" + c + "$" + strconv.Itoa(rows)
	// fmt.Println(areas)
	// f.SetDefinedName(&excelize.DefinedName{
	// 	Name:     "_xlnm.Print_Area",
	// 	RefersTo: firstSheet + "!$A$1:$" + c + "$" + strconv.Itoa(rows),
	// 	Scope:    "Sheet1",
	// })
}
