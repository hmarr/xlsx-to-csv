package main

import (
	"encoding/csv"
	"errors"
	"fmt"
	"os"

	"github.com/jessevdk/go-flags"
	"github.com/xuri/excelize/v2"
)

type Options struct {
	ListSheets bool   `short:"l" long:"list-sheets" description:"List sheets"`
	Sheet      string `short:"s" long:"sheet" description:"Sheet to convert"`
	InputFile  string `short:"i" long:"input" description:"Input XLSX file (default: stdin)"`
	OutputFile string `short:"o" long:"output" description:"Output CSV file (default: stdout)"`
}

var options Options

func main() {
	if _, err := flags.Parse(&options); err != nil {
		switch flagsErr := err.(type) {
		case flags.ErrorType:
			if flagsErr == flags.ErrHelp {
				os.Exit(0)
			}
			os.Exit(1)
		default:
			os.Exit(1)
		}
	}

	if err := runCli(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}

func runCli() error {
	inFile := os.Stdin
	if options.InputFile != "" {
		var err error
		inFile, err = os.Open(options.InputFile)
		if err != nil {
			return err
		}
		defer inFile.Close()
	}

	f, err := excelize.OpenReader(inFile)
	if err != nil {
		return err
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Fprintln(os.Stderr, err)
		}
	}()

	if options.ListSheets {
		listSheets(f)
		return nil
	}

	return sheetToCSV(f)
}

func listSheets(f *excelize.File) {
	for _, sheet := range f.GetSheetList() {
		fmt.Println(sheet)
	}
}

func sheetToCSV(f *excelize.File) error {
	if options.Sheet == "" {
		sheetNames := f.GetSheetList()
		if len(sheetNames) == 0 {
			return errors.New("no sheets found in file")
		}

		options.Sheet = sheetNames[0]
		fmt.Fprintf(os.Stderr, "no sheet specified, using '%s'\n", options.Sheet)
	}

	rows, err := f.GetRows(options.Sheet)
	if err != nil {
		if _, ok := err.(excelize.ErrSheetNotExist); ok {
			return fmt.Errorf("sheet '%s' does not exist", options.Sheet)
		}
		return err
	}

	outFile := os.Stdout
	if options.OutputFile != "" {
		outFile, err = os.Create(options.OutputFile)
		if err != nil {
			return err
		}
		defer outFile.Close()
	}
	out := csv.NewWriter(outFile)
	defer out.Flush()

	for _, row := range rows {
		if err := out.Write(row); err != nil {
			return err
		}
	}

	return nil
}
