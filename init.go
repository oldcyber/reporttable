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
		fmt.Println("Ошибка: ", err)
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
	var realRows int = 0

	fmt.Println("Строк: ", totalRows, "Колонок: ", totalCols)

	// Стили
	var (
		top    = excelize.Border{Type: "top", Style: 1, Color: "000000"}
		left   = excelize.Border{Type: "left", Style: 1, Color: "000000"}
		right  = excelize.Border{Type: "right", Style: 1, Color: "000000"}
		bottom = excelize.Border{Type: "bottom", Style: 1, Color: "000000"}
	)
	// --- Заголовок ---
	styleLightGreyCenterWrap, err := f.NewStyle(&excelize.Style{
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#D9D9D9"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
		Border:    []excelize.Border{top, left, right, bottom},
		Font:      &excelize.Font{Bold: true}})
	if err != nil {
		fmt.Println(err)
	}
	// styleLightGreyLeftWrap, err := f.NewStyle(&excelize.Style{
	// 	Fill:      excelize.Fill{Type: "pattern", Color: []string{"#D9D9D9"}, Pattern: 1},
	// 	Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center", WrapText: true},
	// 	Border:    []excelize.Border{top, left, right, bottom}})
	// if err != nil {
	// 	fmt.Println(err)
	// }
	// styleLightGreyLeftNoWrap, err := f.NewStyle(&excelize.Style{
	// 	Fill:      excelize.Fill{Type: "pattern", Color: []string{"#D9D9D9"}, Pattern: 1},
	// 	Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
	// 	Border:    []excelize.Border{top, left, right, bottom}})
	// if err != nil {
	// 	fmt.Println(err)
	// }
	//--- Простые ячейки ---
	styleColumnCenterWrap, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
		Border:    []excelize.Border{top, left, right, bottom}})
	if err != nil {
		fmt.Println(err)
	}
	styleColumnLeftWrap, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center", WrapText: true},
		Border:    []excelize.Border{top, left, right, bottom}})
	if err != nil {
		fmt.Println(err)
	}
	styleColumnLeftNoWrap, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
		Border:    []excelize.Border{top, left, right, bottom}})
	if err != nil {
		fmt.Println(err)
	}
	// --- Группировка первого уровня ---
	styleGreyCenterWrap, err := f.NewStyle(&excelize.Style{
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#BFBFBF"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
		Border:    []excelize.Border{top, left, right, bottom}})
	if err != nil {
		fmt.Println(err)
	}
	styleGreyLeftWrap, err := f.NewStyle(&excelize.Style{
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#BFBFBF"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center", WrapText: true},
		Border:    []excelize.Border{top, left, right, bottom}})
	if err != nil {
		fmt.Println(err)
	}
	styleGreyLeftNoWrap, err := f.NewStyle(&excelize.Style{
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#BFBFBF"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
		Border:    []excelize.Border{top, left, right, bottom}})
	if err != nil {
		fmt.Println(err)
	}
	// --- Группировка второго уровня ---
	styleBlueCenterWrap, err := f.NewStyle(&excelize.Style{
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#95B3D7"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
		Border:    []excelize.Border{top, left, right, bottom}})
	if err != nil {
		fmt.Println(err)
	}
	styleBlueLeftWrap, err := f.NewStyle(&excelize.Style{
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#95B3D7"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center", WrapText: true},
		Border:    []excelize.Border{top, left, right, bottom}})
	if err != nil {
		fmt.Println(err)
	}
	styleBlueLeftNoWrap, err := f.NewStyle(&excelize.Style{
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#95B3D7"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
		Border:    []excelize.Border{top, left, right, bottom}})
	if err != nil {
		fmt.Println(err)
	}
	// Шапка таблицы
	styleLightBlueCenterWrap, err := f.NewStyle(&excelize.Style{
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#DEE6F0"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
		Border:    []excelize.Border{top, left, right, bottom},
		Font:      &excelize.Font{Bold: true}})
	if err != nil {
		fmt.Println(err)
	}

	// Форматируем документ
	start1 := time.Now()
	// setDocumentProperties(f, firstSheet)
	// ----------------------------------------------------------------

	// Обрезаем первый столбец и первую строку
	err = f.RemoveCol(firstSheet, "A")
	if err != nil {
		fmt.Println(err)
	}
	err = f.RemoveRow(firstSheet, 1)
	if err != nil {
		fmt.Println(err)
	}
	// err = f.RemoveCol(firstSheet, "AC")

	// Скрываем колонку А, в соответствии с заданием
	if err := f.SetColVisible(firstSheet, "A", false); err != nil {
		fmt.Println(err)
	}
	// Общие настройки
	if err := f.SetSheetPrOptions(firstSheet,
		excelize.FitToPage(true),
		excelize.AutoPageBreaks(false),
		excelize.OutlineSummaryBelow(false),
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

	// fmt.Println((intResult))

	// Нижний колонтитул
	if err := f.SetHeaderFooter(firstSheet, &excelize.FormatHeaderFooter{
		// EvenFooter: `&C &"Страница &"&P &"из &"&N`,
		EvenFooter: `&CСтраница &P из &N`,
		OddFooter:  `&CСтраница &P из &N`,
	}); err != nil {
		fmt.Println(err)
	}
	// Верхняя группировка двух последних столбцов
	err = f.SetColOutlineLevel("Sheet1", "AA", 2)
	if err != nil {
		fmt.Println(err)
	}
	err = f.SetColOutlineLevel("Sheet1", "AB", 2)
	if err != nil {
		fmt.Println(err)
	}
	// Скрытие двух последних столбцов
	if err := f.SetColVisible(firstSheet, "AA", false); err != nil {
		fmt.Println(err)
	}
	if err := f.SetColVisible(firstSheet, "AB", false); err != nil {
		fmt.Println(err)
	}
	// Общее оформление таблицы

	// err = f.SetCellStyle(firstSheet, "B"+strconv.Itoa(intResult[0]-1), "D"+strconv.Itoa(intResult[0]-1), styleLightGreyCenterWrap)
	// err = f.SetCellStyle(firstSheet, "F"+strconv.Itoa(intResult[0]-1), "Z"+strconv.Itoa(intResult[0]-1), styleLightGreyCenterWrap)
	// err = f.SetCellStyle(firstSheet, "E"+strconv.Itoa(intResult[0]-1), "E"+strconv.Itoa(intResult[0]-1), styleLightGreyLeftWrap)
	// err = f.SetCellStyle(firstSheet, "AA"+strconv.Itoa(intResult[0]-1), "AA"+strconv.Itoa(intResult[0]-1), styleLightGreyLeftNoWrap)

	// err = f.SetCellStyle(firstSheet, "A"+strconv.Itoa(intResult[0]-1), c+strconv.Itoa(intResult[0]-1), styleLightGrey)
	// err = f.SetCellStyle(firstSheet, "A"+strconv.Itoa(intResult[0]-1), c+strconv.Itoa(intResult[0]-1), styleColumnCenterWrap)
	if err != nil {
		fmt.Println(err)
	}

	// Ширина колонок и выравнивание текста по столбцам
	colWidth := []float64{15, 11, 14.5, 45.67, 21.5, 7.5, 8.83, 11.17, 11.83, 6.5, 6.5, 6.5, 6.5, 6.17, 6.17, 6.17, 6.17, 6.17, 6.17, 6.17, 6.17, 9.17, 10.83, 7.83, 10.67, 11.17, 8}

	for i := 0; i < len(colWidth); i++ {
		c, err := excelize.ColumnNumberToName(i + 2)
		if err != nil {
			fmt.Println(err)
		}
		if err := f.SetColWidth(firstSheet, c, c, colWidth[i]+0.83); err != nil {
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
		// if a <= 2 {
		// 	for i := 0; i < len(intResult); i++ {

		// 		lib.SetRowsColor(firstSheet, intResult[i], totalCols, a, f)
		// 	}

		// }

		// Число строк в таблице с данными
		if len(intResult) > 0 {
			if realRows < intResult[(len(intResult))-1] {
				realRows = intResult[(len(intResult))-1]
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
	// автофильтрация
	result, err := f.SearchSheet(firstSheet, "^\\d*\\.$", true)
	if err != nil {
		fmt.Println(err)
	}
	intResult := lib.ConvStrInt(result)

	c, err := excelize.ColumnNumberToName(totalCols - 1)
	if err != nil {
		fmt.Println(err)
	}
	err = f.AutoFilter(firstSheet, "B"+strconv.Itoa(intResult[0]-1), c+strconv.Itoa(intResult[0]-1), "")
	if err != nil {
		fmt.Println(err)
	}

	fmt.Println("B" + strconv.Itoa(intResult[0]-1))
	err = f.SetCellStyle(firstSheet, "B"+strconv.Itoa(intResult[0]-1), c+strconv.Itoa(intResult[0]-1), styleLightGreyCenterWrap)
	if err != nil {
		fmt.Println(err)

	}

	// Общее форматирование таблицы
	fmt.Println("Строк в таблице: ", realRows)

	for i := 0; i < realRows; i++ {
		err = f.SetCellStyle(firstSheet, "B"+strconv.Itoa(intResult[0]), "D"+strconv.Itoa(realRows), styleColumnCenterWrap)
		if err != nil {
			fmt.Println(err)
			break
		}
		err = f.SetCellStyle(firstSheet, "F"+strconv.Itoa(intResult[0]), "AA"+strconv.Itoa(realRows), styleColumnCenterWrap)
		if err != nil {
			fmt.Println(err)
			break
		}
		err = f.SetCellStyle(firstSheet, "E"+strconv.Itoa(intResult[0]), "E"+strconv.Itoa(realRows), styleColumnLeftWrap)
		if err != nil {
			fmt.Println(err)
			break
		}
		err = f.SetCellStyle(firstSheet, "AB"+strconv.Itoa(intResult[0]), "AB"+strconv.Itoa(realRows), styleColumnLeftNoWrap)
		if err != nil {
			fmt.Println(err)
			break
		}
	}
	// Красим группы
	for a := 1; a < 3; a++ {
		result, err := f.SearchSheet(firstSheet, searchPrefix+strings.Repeat(searchPattern, a)+searchPostfix, true)
		if err != nil {
			fmt.Println(err)
			break
		}
		// Конвертируем результат
		intResult := lib.ConvStrInt(result)
		if a == 1 {
			for i := 0; i < len(intResult); i++ {
				err = f.SetCellStyle(firstSheet, "B"+strconv.Itoa(intResult[i]), "D"+strconv.Itoa(intResult[i]), styleGreyCenterWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
				err = f.SetCellStyle(firstSheet, "F"+strconv.Itoa(intResult[i]), "AA"+strconv.Itoa(intResult[i]), styleGreyCenterWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
				err = f.SetCellStyle(firstSheet, "E"+strconv.Itoa(intResult[i]), "E"+strconv.Itoa(intResult[i]), styleGreyLeftWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
				err = f.SetCellStyle(firstSheet, "AB"+strconv.Itoa(intResult[i]), "AB"+strconv.Itoa(intResult[i]), styleGreyLeftNoWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
			}

		} else if a == 2 {
			for i := 0; i < len(intResult); i++ {
				err = f.SetCellStyle(firstSheet, "B"+strconv.Itoa(intResult[i]), "D"+strconv.Itoa(intResult[i]), styleBlueCenterWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
				err = f.SetCellStyle(firstSheet, "F"+strconv.Itoa(intResult[i]), "AA"+strconv.Itoa(intResult[i]), styleBlueCenterWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
				err = f.SetCellStyle(firstSheet, "E"+strconv.Itoa(intResult[i]), "E"+strconv.Itoa(intResult[i]), styleBlueLeftWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
				err = f.SetCellStyle(firstSheet, "AB"+strconv.Itoa(intResult[i]), "AB"+strconv.Itoa(intResult[i]), styleBlueLeftNoWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
			}
		}

	}

	//Оформляем оглавление таблицы
	/*
		Записать в массив ячейки первой второй и третьей строки. Если в соседней ячейке есть значение != "" Значит объединяем вниз, в противном случае вправо.
		B-F объединение 3 строк
	*/
	firstHead := []string{}
	secondHead := []string{}
	thirdHead := []string{}

	for q := 0; q <= 2; q++ {
		for i := 2; i < totalCols; i++ {
			c, err = excelize.ColumnNumberToName(i)
			if err != nil {
				fmt.Println(err)
			}

			if q == 0 {
				cellValue, err := f.GetCellValue(firstSheet, c+strconv.Itoa(intResult[0]-4))
				if err != nil {
					fmt.Println(err)
				}
				firstHead = append(firstHead, cellValue)
			} else if q == 1 {
				cellValue, err := f.GetCellValue(firstSheet, c+strconv.Itoa(intResult[0]-3))
				if err != nil {
					fmt.Println(err)
				}
				secondHead = append(secondHead, cellValue)
			} else if q == 2 {
				cellValue, err := f.GetCellValue(firstSheet, c+strconv.Itoa(intResult[0]-2))
				if err != nil {
					fmt.Println(err)
				}
				thirdHead = append(thirdHead, cellValue)
			}

		}
	}

	mergeCell := 0

	// Первый уровень
	for i := 0; i < totalCols-2; i++ {
		c, err = excelize.ColumnNumberToName(i + 2)
		if err != nil {
			fmt.Println(err)
		}
		if i < len(firstHead)-1 {
			if firstHead[i] != "" && firstHead[i+1] != "" {
				if err = f.MergeCell(firstSheet, c+strconv.Itoa(intResult[0]-4), c+strconv.Itoa(intResult[0]-2)); err != nil {
					fmt.Println(err)
					return
				}
				err = f.SetCellStyle(firstSheet, c+strconv.Itoa(intResult[0]-4), c+strconv.Itoa(intResult[0]-2), styleLightBlueCenterWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
			} else if firstHead[i] != "" && firstHead[i+1] == "" {
				for x := i + 1; firstHead[x] == ""; x++ {
					mergeCell = x
				}
				v, err := excelize.ColumnNumberToName(i + 2)
				if err != nil {
					fmt.Println(err)
				}
				b, err := excelize.ColumnNumberToName(mergeCell + 2)
				if err != nil {
					fmt.Println(err)
				}
				if err = f.MergeCell(firstSheet, v+strconv.Itoa(intResult[0]-4), b+strconv.Itoa(intResult[0]-4)); err != nil {
					fmt.Println(err)
					return
				}
				err = f.SetCellStyle(firstSheet, v+strconv.Itoa(intResult[0]-4), b+strconv.Itoa(intResult[0]-4), styleLightBlueCenterWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
			}
		} else if i < len(firstHead) {
			if firstHead[i] != "" {
				if err = f.MergeCell(firstSheet, c+strconv.Itoa(intResult[0]-4), c+strconv.Itoa(intResult[0]-2)); err != nil {
					fmt.Println(err)
					return
				}
				err = f.SetCellStyle(firstSheet, c+strconv.Itoa(intResult[0]-4), c+strconv.Itoa(intResult[0]-2), styleLightBlueCenterWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
			}
		}
	}
	//Второй уровень
	for i := 0; i < totalCols-2; i++ {
		c, err = excelize.ColumnNumberToName(i + 2)
		if err != nil {
			fmt.Println(err)
		}
		if i < len(secondHead)-1 {
			if secondHead[i] != "" && secondHead[i+1] != "" {
				if err = f.MergeCell(firstSheet, c+strconv.Itoa(intResult[0]-3), c+strconv.Itoa(intResult[0]-3)); err != nil {
					fmt.Println(err)
					return
				}
				err = f.SetCellStyle(firstSheet, c+strconv.Itoa(intResult[0]-3), c+strconv.Itoa(intResult[0]-3), styleLightBlueCenterWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
			} else if secondHead[i] != "" && secondHead[i+1] == "" {
				for x := i + 1; secondHead[x] == "" && x < len(secondHead)-1 && firstHead[x] == ""; x++ {
					mergeCell = x
				}
				v, err := excelize.ColumnNumberToName(i + 2)
				if err != nil {
					fmt.Println(err)
				}
				b, err := excelize.ColumnNumberToName(mergeCell + 2)
				if err != nil {
					fmt.Println(err)
				}
				if err = f.MergeCell(firstSheet, v+strconv.Itoa(intResult[0]-3), b+strconv.Itoa(intResult[0]-3)); err != nil {
					fmt.Println(err)
					return
				}
				err = f.SetCellStyle(firstSheet, v+strconv.Itoa(intResult[0]-3), b+strconv.Itoa(intResult[0]-3), styleLightBlueCenterWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
			}
		}
	}
	//Третий уровень
	for i := 0; i < totalCols-2; i++ {
		c, err = excelize.ColumnNumberToName(i + 2)
		if err != nil {
			fmt.Println(err)
		}
		if i < len(thirdHead)-1 {
			if thirdHead[i] != "" && thirdHead[i+1] != "" {
				if err = f.MergeCell(firstSheet, c+strconv.Itoa(intResult[0]-2), c+strconv.Itoa(intResult[0]-2)); err != nil {
					fmt.Println(err)
					return
				}
				err = f.SetCellStyle(firstSheet, c+strconv.Itoa(intResult[0]-2), c+strconv.Itoa(intResult[0]-2), styleLightBlueCenterWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
			} else if thirdHead[i] != "" && thirdHead[i+1] == "" {
				for x := i + 1; thirdHead[x] == "" && x < len(thirdHead)-1 && firstHead[x] == ""; x++ {
					mergeCell = x
				}
				v, err := excelize.ColumnNumberToName(i + 2)
				if err != nil {
					fmt.Println(err)
				}
				b, err := excelize.ColumnNumberToName(mergeCell + 2)
				if err != nil {
					fmt.Println(err)
				}
				if err = f.MergeCell(firstSheet, v+strconv.Itoa(intResult[0]-2), b+strconv.Itoa(intResult[0]-2)); err != nil {
					fmt.Println(err)
					return
				}
				err = f.SetCellStyle(firstSheet, v+strconv.Itoa(intResult[0]-2), b+strconv.Itoa(intResult[0]-2), styleLightBlueCenterWrap)
				if err != nil {
					fmt.Println(err)
					break
				}
			}
		}
	}
	//Высота строк в шапке таблицы
	if err := f.SetRowHeight(firstSheet, intResult[0]-3, 96+0.73); err != nil {
		fmt.Println(err)
	}

	// ---
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
