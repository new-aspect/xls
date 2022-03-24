package xls

import (
	"fmt"
	"testing"
)

func TestExampleOpen(t *testing.T) {
	if xlFile, err := Open("Table.xls", "utf-8"); err == nil {
		fmt.Println(xlFile.Author)
	}
}

func TestExampleWorkBook_NumberSheets(t *testing.T) {
	if xlFile, err := Open("Table.xls", "utf-8"); err == nil {
		for i := 0; i < xlFile.NumSheets(); i++ {
			sheet := xlFile.GetSheet(i)
			fmt.Println(sheet.Name)
		}
	}
}

//Output: read the content of first two cols in each row
func TestExampleWorkBook_GetSheet(t *testing.T) {
	if xlFile, err := Open("Table.xls", "utf-8"); err == nil {
		if sheet1 := xlFile.GetSheet(0); sheet1 != nil {
			fmt.Print("Total Lines ", sheet1.MaxRow, sheet1.Name)
			col1 := sheet1.Row(0).Col(0)
			col2 := sheet1.Row(0).Col(0)
			for i := 0; i <= (int(sheet1.MaxRow)); i++ {
				row1 := sheet1.Row(i)
				col1 = row1.Col(0)
				col2 = row1.Col(1)
				fmt.Print("\n", col1, ",", col2)
			}
		}
	}
}

func TestExampleWorkBook_GetAll(t *testing.T) {
	fileName := "Table.xls"
	xlFile, err := Open(fileName, "utf-8")
	if err != nil {
		panic(err)
		return
	}

	for k := 0; k < xlFile.NumSheets(); k++ {
		if sheet1 := xlFile.GetSheet(k); sheet1 != nil {
			fmt.Print("Total Lines ", sheet1.MaxRow, sheet1.Name)
			for i := 0; i <= (int(sheet1.MaxRow)); i++ {
				row1 := sheet1.Row(i)
				for j := 0; j < row1.LastCol(); j++ {
					fmt.Print(row1.Col(j) + " ")
				}
				fmt.Print("/n")
			}
		}
	}

}
