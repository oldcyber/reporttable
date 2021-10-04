//package reporttable
package main

import (
	"flag"
	"fmt"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
	"oldcyber.xyz/reporttable/lib"
)

func main() {
	// Запускаем таймер
	fmt.Println("Starting... ")
	start := time.Now()

	inpPath := flag.String("i", "", "input path")
	outPath := flag.String("o", "", "output path")
	const pageLines int = 100

	flag.Parse()

	// inpPath := "input.xlsx"
	// outPath := "result.xlsx"

	f, err := excelize.OpenFile(*inpPath)
	if err != nil {
		fmt.Println(err)
		return
	}

	// Записываем имя активного листа
	firstSheet := f.WorkBook.Sheets.Sheet[0].Name
	// Количество строк
	rows, err := f.GetRows(firstSheet)
	if err != nil {
		fmt.Println(err)
	}
	// Количество колонок
	cols, err := f.GetCols(firstSheet)
	if err != nil {
		fmt.Println(err)
	}
	var totalRows int = len(rows)
	var totalCols int = len(cols)

	// fmt.Println("Строк: ", totalRows, "Колонок: ", totalCols)

	// Форматируем документ
	start1 := time.Now()
	// setDocumentProperties(f, firstSheet)
	// ----------------------------------------------------------------

	// Скрываем колонку А, в соответствии с заданием
	if err := f.SetColVisible(firstSheet, "A", false); err != nil {
		fmt.Println(err)
	}
	// Общие настройки
	if err := f.SetSheetPrOptions(firstSheet,
		excelize.FitToPage(true),
		excelize.AutoPageBreaks(false),
	); err != nil {
		fmt.Println(err)
	}

	// Формат документа
	if err := f.SetPageLayout(
		firstSheet,
		excelize.BlackAndWhite(false),
		excelize.FirstPageNumber(1),
		excelize.PageLayoutOrientation(excelize.OrientationLandscape),
		excelize.PageLayoutPaperSize(8),
		excelize.FitToHeight(10000),
		excelize.FitToWidth(1),
		excelize.PageLayoutScale(50),
	); err != nil {
		fmt.Println(err)
	}
	// Поля
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

	result, err := f.SearchSheet(firstSheet, "^\\d*\\.$", true)
	if err != nil {
		fmt.Println(err)
	}
	intResult := lib.ConvStrInt(result)
	// fmt.Println((intResult))

	// автофильтрация
	c, err := excelize.ColumnNumberToName(totalCols)
	if err != nil {
		fmt.Println(err)
	}
	err = f.AutoFilter(firstSheet, "A"+strconv.Itoa(intResult[0]-1), c+strconv.Itoa(intResult[0]-1), "")
	if err != nil {
		fmt.Println(err)
	}
	styleLightGrey, err := f.NewStyle(`{"fill":{"type":"pattern","color":["#D9D9D9"],"pattern":1}}`)
	if err != nil {
		fmt.Println(err)
	}
	err = f.SetCellStyle(firstSheet, "A"+strconv.Itoa(intResult[0]-1), c+strconv.Itoa(intResult[0]-1), styleLightGrey)
	if err != nil {
		fmt.Println(err)
	}

	// Нижний колонтитул
	if err := f.SetHeaderFooter(firstSheet, &excelize.FormatHeaderFooter{
		// EvenFooter: `&C &"Страница &"&P &"из &"&N`,
		EvenFooter: `&CСтраница &P из &N`,
		OddFooter:  `&CСтраница &P из &N`,
	}); err != nil {
		fmt.Println(err)
	}

	// Ширина колонок
	colWidth := []float64{15.83, 11.83, 15.33, 46.5, 22.33, 8.33, 9.67, 12, 12.67, 7.33, 7.33, 7.33, 7.33, 7, 7, 7, 7, 7, 7, 7, 7, 10, 11.67, 8.67, 8.67, 11.67, 59.5}

	for i := 0; i < len(colWidth); i++ {
		c, err := excelize.ColumnNumberToName(i + 2)
		if err != nil {
			fmt.Println(err)
		}
		// fmt.Println("c: ", c, "width: ", colWidth[i])
		if err := f.SetColWidth(firstSheet, c, c, colWidth[i]); err != nil {
			fmt.Println(err)
		}

	}
	duration := time.Since(start1)
	fmt.Println("Page config: ", duration.String())
	// ----------------------------------------------------------------
	// Группируем, в соответствии со значениями в колонке В
	start2 := time.Now()
	// Ищем по regexp все варианты, пока не будет 0

	var searchPattern string = "\\d*\\."
	var searchPrefix string = "^"
	var searchPostfix string = "$"
	var groups int = 7

	for a := 1; a < groups; a++ {
		result, err := f.SearchSheet(firstSheet, searchPrefix+strings.Repeat(searchPattern, a)+searchPostfix, true)
		if err != nil {
			fmt.Println(err)
			break
		}

		// Конвертируем результат
		intResult := lib.ConvStrInt(result)
		// Красим
		if a <= 2 {
			for i := 0; i < len(intResult); i++ {
				lib.SetRowsColor(firstSheet, intResult[i], totalCols, a, f)
			}

		}

		// Расставляем страницы
		start3 := time.Now()

		var wrongLine bool = false
		if a == 1 {
			for i := pageLines; i < totalRows; {
				for _, v := range intResult {
					if v == i {
						wrongLine = true
					} else {
						wrongLine = false
					}
				}

				if wrongLine {
					// fmt.Println("неправильный A" + strconv.Itoa(i-1))
					if err := f.InsertPageBreak(firstSheet, "A"+strconv.Itoa(i-1)); err != nil {
						fmt.Println(err)
					}
				} else {
					// fmt.Println("A" + strconv.Itoa(i))
					if err := f.InsertPageBreak(firstSheet, "A"+strconv.Itoa(i)); err != nil {
						fmt.Println(err)
					}
				}

				i = i + pageLines
			}
			duration = time.Since(start3)
			fmt.Println("Page breaks: ", duration.String())

		}

		// Добавляем группы
		for b := 0; b < len(intResult); b++ {

			if b+1 < len(intResult) {
				for r := intResult[b] + 1; r <= intResult[b+1]-1; r++ {

					if intResult[b+1]-1 <= intResult[b]+1 {
					} else {

						if err := f.SetRowOutlineLevel(firstSheet, r, uint8(a)); err != nil {
							fmt.Println(err)
						}
					}

				}
			} else if len(intResult)%2 != 0 {
				for r := intResult[b] + 1; r <= len(rows); r++ {

					if err := f.SetRowOutlineLevel(firstSheet, r, uint8(a)); err != nil {
						fmt.Println(err)
					}
				}
			}
		}
	}

	// Закрываем группы (в обратном порядке)
	for a := groups; a > 0; a-- {

		result, err := f.SearchSheet(firstSheet, searchPrefix+strings.Repeat(searchPattern, a)+searchPostfix, true)

		if err != nil {
			fmt.Println(err)
			break
		}
		// Конвертируем результат
		intResult := lib.ConvStrInt(result)
		// Закрываем группу

		for i := range intResult {
			if err := f.SetRowOutlineLevel(firstSheet, intResult[i], uint8(a)); err != nil {
				fmt.Println(err)
			}

		}

	}
	duration = time.Since(start2)
	fmt.Println("Grouping: ", duration.String())
	//сохраняем результат
	if err = f.SaveAs(*outPath); err != nil {
		println(err.Error())
	}
	//Печать результат выполнения
	duration = time.Since(start)
	fmt.Println("Total time: ", duration.String())
	fmt.Println(" Done! ")
}
