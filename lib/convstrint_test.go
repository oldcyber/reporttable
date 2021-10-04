package lib

import (
	"testing"
)

func TestConvStrInt(t *testing.T) {
	tables := []struct {
		desc string   //Описание кейса тестирования
		in   []string // Входные данные
		out  []int    // Выходные данные
	}{
		{"Правильные данные",
			[]string{"B1", "B2", "B3", "B4", "B22", "B29", "B333", "B399", "B4444", "B4999", "B55555", "B59999"},
			[]int{1, 2, 3, 4, 22, 29, 333, 399, 4444, 4999, 55555, 59999},
		},
	}

	for i, table := range tables {
		res := ConvStrInt(table.in)

		if table.out[i] != res[i] {
			t.Error("Error. Input and Output mismatch", res[i], table.out[i])
		}
	}
}
