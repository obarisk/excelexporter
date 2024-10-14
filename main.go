package main

import (
	"errors"
	"flag"
	"fmt"
	"io"
	"log/slog"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

var (
	errArgParse = errors.New("argument parsing error")
	slgr        = slog.New(slog.NewTextHandler(os.Stderr, nil))
)

type Args struct {
	InputFile    string
	OutputFile   string
	SheetName    string
	RowsToRemove uint
}

func readArgs() (Args, error) {
	excelFile := flag.String("f", "", "Path to the Excel file")
	lineToRemove := flag.Int("l", 0, "Line number to remove")
	sheetName := flag.String("s", "", "Sheet name")
	newFileName := flag.String("o", "", "New file name")

	flag.Parse()

	arg := Args{
		InputFile:    *excelFile,
		SheetName:    *sheetName,
		RowsToRemove: uint(*lineToRemove),
		OutputFile:   *newFileName,
	}

	if arg.InputFile == "" || !strings.HasSuffix(arg.InputFile, ".xlsx") || arg.RowsToRemove < 0 {
		return Args{}, errArgParse
	}

	if _, err := os.Stat(arg.InputFile); os.IsNotExist(err) {
		return Args{}, err
	}

	if arg.SheetName == "" {
		slgr.Info("no sheet name provided, using sheet with index = 1 as default")
	}

	if arg.OutputFile == "" {
		arg.OutputFile = strings.TrimSuffix(arg.InputFile, ".xlsx") + "_cleaned.xlsx"
	}
	p := arg.OutputFile
	for i := 0; ; i++ {
		if _, err := os.Stat(p); os.IsNotExist(err) {
			break
		}
		p = strings.TrimSuffix(p, ".xlsx") + "_" + strconv.Itoa(i) + ".xlsx"
	}
	arg.OutputFile = p
	return arg, nil
}

func copy(src, dst string) (int64, error) {
	sourceFileStat, err := os.Stat(src)
	if err != nil {
		return 0, err
	}

	if !sourceFileStat.Mode().IsRegular() {
		return 0, fmt.Errorf("%s is not a regular file", src)
	}

	source, err := os.Open(src)
	if err != nil {
		return 0, err
	}
	defer source.Close()

	destination, err := os.Create(dst)
	if err != nil {
		return 0, err
	}
	defer destination.Close()
	nBytes, err := io.Copy(destination, source)
	return nBytes, err
}

func main() {
	args, err := readArgs()
	if err != nil {
		slgr.Info(`excelexporter -f file -l rows [-s sheet -o output]`)
		slgr.Error(err.Error())
		os.Exit(1)
	}
	_, err = copy(args.InputFile, args.OutputFile)
	if err != nil {
		slgr.Error(err.Error())
		os.Exit(1)
	}

	xlsx, err := excelize.OpenFile(args.OutputFile)
	if err != nil {
		slgr.Error(err.Error())
		os.Exit(1)
	}
	defer xlsx.Close()

	sht := args.SheetName
	if sht == "" {
		sht = xlsx.GetSheetName(0)
	}

	for i := 0; i < int(args.RowsToRemove); i++ {
		if err := xlsx.RemoveRow(sht, 1); err != nil {
			slgr.Error(err.Error())
			os.Exit(1)
		}
	}
	xlsx.Save()
}
