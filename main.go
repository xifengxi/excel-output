package main

import (
	"flag"
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	flag.Parse()
	dir := flag.Args()
	if len(dir) == 0 {
		fmt.Println(fmt.Sprintf("DirRequired"))
		return
	}
	getExcelData(dir[0])
}

type ExcelInfo struct {
	which   []string
	titles  []string
	titles2 []string
	lie     string
	hang    string
	name    string
}

func getExcelData(dir string) {
	f, err := excelize.OpenFile(dir)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	for _, name := range f.GetSheetList() {
		//fmt.Println("SheetName: " + name)
		if strings.HasPrefix(name, "!") {
			continue
		}

		isRow := true // 默认是行表
		if strings.HasPrefix(name, "~") {
			isRow = false
		}

		titleObj := &ExcelInfo{
			name: name,
		}
		idx := 0
		if isRow {
			titleObj.hang = "行"
			titleObj.lie = "列"
			row, _ := f.Rows(name)
			for row.Next() {
				col, _ := row.Columns()
				printData(idx, col, titleObj)
				idx++
			}
		} else {
			row, _ := f.Rows(name)
			for row.Next() {
				data, _ := row.Columns()
				if len(data) >= 5 && !isEmpty(data[0]) {
					data = data[:5]
					fmt.Println(fmt.Sprintf("[%s]表[%s]行=%s", name, data[2], data))
				}
			}
		}
		fmt.Println("")

		//WriterCSV := csv.NewWriter(os.Stdout)

		//rows, err := f.Rows(name)
		//if err != nil {
		//	fmt.Println(err)
		//	return
		//}
		//rowCheck := isRow
		//rowNum := 0
		//line := 0
		//for rows.Next() {
		//	line++
		//	fmt.Println("" + name + "_Line:")
		//	row, err := rows.Columns()
		//	if err != nil {
		//		fmt.Println(err)
		//	}
		//	if isRow {
		//		if rowCheck {
		//			rowCheck = false
		//			for i, s := range row {
		//				rowNum = i
		//				if isEmpty(s) {
		//					break
		//				}
		//			}
		//		}
		//	} else {
		//		// 如果是列表，那么看一下第一个是否为none或者空
		//		if len(row) == 0 || isEmpty(row[0]) {
		//			continue
		//		}
		//	}
		//	if rowNum > 0 && len(row) >= rowNum {
		//		row = row[:rowNum]
		//	}
		//	for _, s := range row {
		//		fmt.Println("\t" + strings.Replace(s, "\n", "↩", -1))
		//	}
		//	//csverr := WriterCSV.Write(row)
		//	//if csverr != nil {
		//	//	return
		//	//}
		//}
		////WriterCSV.Flush()
	}
}

func isEmpty(s string) bool {
	r := strings.Trim(s, " ")
	return r == "" || r == "none"
}

func printData(idx int, data []string, info *ExcelInfo) {
	if idx == 0 {
		info.which = data
	} else if idx == 2 {
		info.titles = data
	} else if idx == 3 {
		info.titles2 = data
	}
	if idx < 4 {
		fmt.Println(fmt.Sprintf("[%s]表:%d列", info.name, idx), data)
	} else {
		key := ""
		for i, s := range data {
			if len(info.which) <= i {
				break
			}
			which := info.which[i]
			if isEmpty(which) {
				continue
			}
			title := safeGet(info.titles, i) + "," + safeGet(info.titles2, i)
			if i == 0 {
				if isEmpty(s) {
					break
				}
				key = s
			} else {
				ss := strings.Replace(s, "\n", "↩", -1)
				fmt.Println(fmt.Sprintf("[%s]表[%s]%s[%s]%s=%s", info.name, key, info.hang, title, info.lie, ss))
			}
		}
	}
}

func safeGet(sl []string, idx int) string {
	if len(sl) <= idx {
		return ""
	} else {
		return sl[idx]
	}
}
