package excel

import (
	"errors"
	"github.com/extrame/xls"
	"strings"
)

type Excel interface {
	Open(fileName string) error
	Close() error
	// xls的sheet从0开始，xlsx的sheet从1开始
	Rows(sheetNum int) ([][]string, error)
	FirstSheet() int
}

func OpenExcel(filePath string) (Excel, error) {
	var err error
	if strings.HasSuffix(filePath, ".xls") {
		t := &Xls{}
		err = t.Open(filePath)
		return t, err
	} else {
		t := &Xlsx{}
		err = t.Open(filePath)
		return t, err
	}
}

type Xls struct {
	file *xls.WorkBook
}

func (x *Xls) Open(fileName string) error {
	var err error
	x.file, err = xls.Open(fileName, "utf-8")
	if err != nil {
		return err
	}
	return nil
}

func (x *Xls) Close() error {
	return nil
}

func (x *Xls) FirstSheet() int {
	return 0
}

// xls的sheet从0开始
func (x *Xls) Rows(sheetNum int) ([][]string, error) {
	if x.file.NumSheets() <= sheetNum {
		return nil, errors.New("sheet not found")
	}
	sheet := x.file.GetSheet(sheetNum)
	rowNum := sheet.MaxRow
	data := make([][]string, rowNum)
	for i := 0; i < int(rowNum); i++ {
		row := sheet.Row(i)
		lastCol := row.LastCol()
		cols := make([]string, lastCol)
		if lastCol > 0 {
			for j := 0; j < lastCol; j++ {
				cols[j] = row.Col(j)
			}
			data[i] = cols
		}

	}
	return data, nil
}

func (x *Xls) SaveToWriter(header []string, data [][]string) error {

	return nil
}
