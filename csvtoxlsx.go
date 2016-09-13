package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"os"

	"github.com/psmithuk/xlsx"
)

func main() {
	var (
		InFileName  string
		OutFileName string
		SheetTitle  string
		ShowHelp    bool
		err         error
	)

	flag.StringVar(&InFileName, "in", "input.csv", "input file name")
	flag.StringVar(&OutFileName, "out", "output.xlsx", "output file name")
	flag.StringVar(&SheetTitle, "title", "Sheet1", "sheet title")
	flag.BoolVar(&ShowHelp, "h", false, "show help")
	flag.Parse()

	if ShowHelp {
		flag.Usage()
		os.Exit(0)
	}
//Buffers
	inStream, outStream := os.Stdin, os.Stdout
	if InFileName != "" {
		if inStream, err = os.Open(InFileName); err != nil {
			fmt.Println(err)
			os.Exit(1)
		}
		defer inStream.Close()
	}

	if OutFileName != "" {
		if outStream, err = os.Create(OutFileName); err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}
		defer inStream.Close()
	}

	inCSV := csv.NewReader(inStream)
	outX := xlsx.NewWorkbookWriter(outStream)
	defer outX.Close()
	var (
		sw *xlsx.SheetWriter
		sh xlsx.Sheet
	)
	firstRow := true
	for {
		row, err := inCSV.Read()
		if err == io.EOF {
			break
		} else if err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}
		if firstRow {
			firstRow = false
			cols := make([]xlsx.Column, len(row))
			for i, _ := range row {
				cols[i] = xlsx.Column{
					Name:  fmt.Sprintf("Col%d", i),
					Width: 15,
				}
			}
			sh = xlsx.NewSheetWithColumns(cols)
			sh.Title = SheetTitle
			sw, _ = outX.NewSheetWriter(&sh)
		}

		xRow := sh.NewRow()
		for i, r := range row {
			xRow.Cells[i] = xlsx.Cell{
				Type:  xlsx.CellTypeInlineString,
				Value: r,
			}
		}
		if err = sw.WriteRows([]xlsx.Row{xRow}); err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}
	}
}
