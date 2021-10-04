package lib

import (
	"fmt"
	"testing"
)

func TestSetRowsColor(t *testing.T) {
	tables := []struct {
		firstSheet string
		rows       int
		cols       int
		level      int
		// f *excelize.File
	}{
		{
			"Sheet1",
			10,
			10,
			1,
			// inp,
		},
	}
	fmt.Println(tables)

	// for i, table := range tables {
	// 	SetRowsColor(table.firstSheet, table.rows, table.cols, table.level, f *excelize.File)

	// 	if table.out[i] != res[i] {
	// 		t.Error("Error. Input and Output mismatch", res[i], table.out[i])
	// 	}
	// }
}
