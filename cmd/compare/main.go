package main

import (
	"fmt"
	"math"
	"os"
	"strconv"
	"strings"

	xls "github.com/nicholasgasior/go-extreme-xls"
	"github.com/tealeg/xlsx"
)

func main() {
	xlsPath := os.Args[1]
	xlsxPath := os.Args[2]
	debug := len(os.Args) > 3 && os.Args[3] == "-debug"

	xlsFile, err := xls.Open(xlsPath, "utf-8")
	if err != nil {
		fmt.Fprintf(os.Stderr, "open xls: %v\n", err)
		os.Exit(1)
	}

	xlsxFile, err := xlsx.OpenFile(xlsxPath)
	if err != nil {
		fmt.Fprintf(os.Stderr, "open xlsx: %v\n", err)
		os.Exit(1)
	}

	diffs := 0
	for sheetIdx, xlsxSheet := range xlsxFile.Sheets {
		xlsSheet, err := xlsFile.GetSheet(sheetIdx)
		if err != nil {
			fmt.Fprintf(os.Stderr, "sheet %d: error: %v\n", sheetIdx, err)
			continue
		}
		if xlsSheet == nil {
			fmt.Fprintf(os.Stderr, "sheet %d: missing in xls\n", sheetIdx)
			continue
		}

		sheetName := xlsxSheet.Name
		for rowIdx, xlsxRow := range xlsxSheet.Rows {
			xlsRow := xlsSheet.Row(rowIdx)
			if xlsRow == nil {
				for colIdx, xlsxCell := range xlsxRow.Cells {
					v := xlsxCell.String()
					if v != "" {
						fmt.Printf("[%s] R%dC%d  xlsx=%q  xls=(missing row)\n", sheetName, rowIdx+1, colIdx+1, v)
						diffs++
					}
				}
				continue
			}

			for colIdx, xlsxCell := range xlsxRow.Cells {
				xlsxText := xlsxCell.String()
				xlsText := xlsRow.Col(colIdx)

				if xlsText == xlsxText {
					continue
				}

				if strings.TrimSpace(xlsText) == "" && strings.TrimSpace(xlsxText) == "" {
					continue
				}

				xlsFloat, xlsErr := strconv.ParseFloat(xlsText, 64)
				xlsxFloat, xlsxErr := strconv.ParseFloat(xlsxText, 64)
				if xlsErr == nil && xlsxErr == nil {
					diff := math.Abs(xlsFloat - xlsxFloat)
					relDiff := float64(0)
					if max := math.Max(math.Abs(xlsFloat), math.Abs(xlsxFloat)); max > 0 {
						relDiff = diff / max
					}
					if relDiff < 1e-9 {
						continue
					}
					fmt.Printf("[%s] R%dC%d  xlsx=%s  xls=%s  (num diff: %e)\n",
						sheetName, rowIdx+1, colIdx+1, xlsxText, xlsText, diff)
				} else {
					fmt.Printf("[%s] R%dC%d  xlsx=%q  xls=%q\n",
						sheetName, rowIdx+1, colIdx+1, xlsxText, xlsText)
				}
				if debug {
					typeName, xfIdx, fmtNo := xlsRow.ColFormatInfo(colIdx)
					fmtStr := ""
					if fmtNo >= 164 {
						if f, ok := xlsFile.Formats[fmtNo]; ok {
							fmtStr = f.String()
						}
					}
					fmt.Printf("    type=%s xfIdx=%d fmtNo=%d fmtStr=%q\n", typeName, xfIdx, fmtNo, fmtStr)
				}
				diffs++
			}
		}
	}

	if diffs == 0 {
		fmt.Println("No differences found.")
	} else {
		fmt.Printf("\n%d difference(s) found.\n", diffs)
	}
}
