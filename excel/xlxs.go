package excel

import (
	"bytes"
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
)

type Xlsx struct {
	file *excelize.File
}

func (x *Xlsx) Open(fileName string) error {
	var err error
	x.file, err = excelize.OpenFile(fileName)
	if err != nil {
		return err
	}
	return nil
}

func (x *Xlsx) Close() error {
	return x.file.Close()
}

func (x *Xlsx) FirstSheet() int {
	return 1
}

// xlsx的sheet从1开始
func (x *Xlsx) Rows(sheetNum int) ([][]string, error) {
	sheetMap := x.file.GetSheetMap()
	sheet, ok := sheetMap[sheetNum]
	if !ok {
		return nil, errors.New("sheet not found")
	}
	rows, err := x.file.GetRows(sheet)
	if err != nil {
		return nil, err
	}
	return rows, nil
}

func (x *Xlsx) NewFile() {
	x.file = excelize.NewFile()
}

func (x *Xlsx) SetSheetData(sheetName string, header []string, data [][]string) error {
	if len(header) < 1 {
		return errors.New("header is empty")
	}
	index := x.file.GetSheetIndex(sheetName)
	if index == -1 {
		index = x.file.NewSheet(sheetName)
	}
	x.file.SetActiveSheet(index)

	// set data
	x.file.SetSheetRow(sheetName, "A1", &header)
	for i := 0; i < len(data); i++ {
		x.file.SetSheetRow(sheetName, fmt.Sprintf("A%d", i+2), &data[i])
	}

	// set style and width
	for i := 0; i < len(header); i++ {
		width := len(header[i])
		if width > 30 {
			width = 30
		}
		end, err := excelize.ColumnNumberToName(i + 1)
		if err != nil {
			return err
		}
		x.file.SetColWidth(sheetName, end, end, float64(width))
	}
	lineStyle, _ := x.file.NewStyle(`{"alignment":{"horizontal":"center","vertical":"center"}}`)
	topStyle, _ := x.file.NewStyle(`{"font":{"bold":true},"alignment":{"horizontal":"center","vertical":"center"}}`)
	x.file.SetRowStyle(sheetName, 1, 1, topStyle)
	if len(data) > 0 {
		x.file.SetRowStyle(sheetName, 2, len(data)+1, lineStyle)
	}
	return nil
}

func (x *Xlsx) SaveAs(filePath string) error {
	return x.file.SaveAs(filePath)
}

func (x *Xlsx) WriteToBuffer() (*bytes.Buffer, error) {
	return x.file.WriteToBuffer()
}

// write the file
func (x *Xlsx) WriteTo(w io.Writer) (int64, error) {
	return x.file.WriteTo(w)
}

func (x *Xlsx) Write(w io.Writer) error {
	return x.file.Write(w)
}
