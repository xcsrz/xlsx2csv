package main

import (
	"encoding/csv"
	"errors"
	"fmt"
	"os"

	"github.com/thedatashed/xlsxreader"
)

func main() {
	filepath := os.Args[1]
	if _, err := os.Stat(filepath); errors.Is(err, os.ErrNotExist) {
		fmt.Println("The file", filepath, "does not exist")
		return
	}
	outputFilepath := filepath + ".csv"

	xls, _ := xlsxreader.OpenFile(filepath)
	defer xls.Close()

	fp, err := os.Create(outputFilepath)
	csvwriter := csv.NewWriter(fp)

	if err != nil {
		panic(err)
	}
	defer fp.Close()
	for row := range xls.ReadRows(xls.Sheets[0]) {
		lastCellIndex := row.Cells[len(row.Cells)-1].ColumnIndex()
		cellValues := make([]string, lastCellIndex+1)
		for _, cell := range row.Cells {
			cellValues[cell.ColumnIndex()] = cell.Value
		}
		_ = csvwriter.Write(cellValues)
	}
	csvwriter.Flush()
}
