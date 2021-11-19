package lib

import (
	"crypto/tls"
	"fmt"
	"log"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
	"gopkg.in/gomail.v2"
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

// Отправка файла по почте
func SendMail(toMail string, fileAttach string) {
	config, err := LoadConfig(".")
	if err != nil {
		log.Fatal("cannot load config:", err)
	}
// TODO:скорректировать тему письма и текст сообщения
	

	d := gomail.NewDialer(config.Server, config.Port, config.Login, config.Password)
	d.TLSConfig = &tls.Config{InsecureSkipVerify: true}

	m := gomail.NewMessage(gomail.SetCharset("UTF-8"))
	m.SetHeader("From", config.From)
	m.SetHeader("To", toMail)
	m.SetHeader("Subject", "Выгрузка готова")
	m.SetBody("text/html", "Отчёт сгенерирован и находится во вложении")
	m.Attach(fileAttach)
	if err := d.DialAndSend(m); err != nil {
		// panic(err)
		fmt.Println(err)
	}

	m.Reset()
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

func SetWorksheetStyle(l int) {

}
