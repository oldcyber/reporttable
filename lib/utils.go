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

func MyPageProperties(firstSheet string, f *excelize.File) {
	// Общие настройки
	if err := f.SetSheetPrOptions(firstSheet,
		excelize.FitToPage(false),
		excelize.AutoPageBreaks(false),
		excelize.OutlineSummaryBelow(false),
	); err != nil {
		fmt.Println(err)
	}
}

func MyPageLayout(firstSheet string, f *excelize.File) {
	if err := f.SetPageLayout(
		firstSheet,
		excelize.BlackAndWhite(false),
		excelize.FirstPageNumber(1),
		excelize.PageLayoutOrientation(excelize.OrientationLandscape),
		excelize.PageLayoutPaperSize(8),
		// excelize.FitToHeight(10000),
		// excelize.FitToWidth(1),
		excelize.PageLayoutScale(60),
	); err != nil {
		fmt.Println(err)
	}
}

func MyPageMargins(firstSheet string, f *excelize.File) {
	if err := f.SetPageMargins(firstSheet,
		excelize.PageMarginBottom(0.75),
		excelize.PageMarginFooter(0.31),
		excelize.PageMarginHeader(0.31),
		excelize.PageMarginLeft(0.79),
		excelize.PageMarginRight(0.39),
		excelize.PageMarginTop(0.47),
	); err != nil {
		fmt.Println(err)
	}
}

// FIXME: Не работает разбивка
func PageBreaks(lines int, firstSheet string, f *excelize.File) {
	rows, err := f.GetRows(firstSheet)
	if err != nil {
		fmt.Println(err)
	}
	var TotalRows int = len(rows)
	for i := 1; i < TotalRows; {
		if err := f.InsertPageBreak(firstSheet, "A"+strconv.Itoa(i)); err != nil {
			fmt.Println(err)
		}
		i = i + lines
	}
}

/* func SetRowsColor(firstSheet string, rows int, cols int, level int, f *excelize.File) {

	// Определяем стили
	styleGrey, err := f.NewStyle(`{"fill":{"type":"pattern","color":[""],"pattern":1}}`)
	if err != nil {
		fmt.Println(err)
	}
	styleBlue, err := f.NewStyle(`{"fill":{"type":"pattern","color":["#95B3D7"],"pattern":1}}`)
	if err != nil {
		fmt.Println(err)
	}

	var firstCell string = "A" + strconv.Itoa(rows)
	c, err := excelize.ColumnNumberToName(cols - 1)
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

	// Область печати
	// var areas string = firstSheet + "!$A$1:$" + c + "$" + strconv.Itoa(rows)
	// fmt.Println(areas)
	// f.SetDefinedName(&excelize.DefinedName{
	// 	Name:     "_xlnm.Print_Area",
	// 	RefersTo: firstSheet + "!$A$1:$" + c + "$" + strconv.Itoa(rows),
	// 	Scope:    "Sheet1",
	// })
}
*/
func SetWorksheetStyle(l int) {

}
