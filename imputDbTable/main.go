package main

import (
	"flag"
	"fmt"
	"os"
	"strconv"
	"strings"

	"../../excelize"
	// "excelize"
)

type mujudaihao struct {
	counts int
	rows   [][]string
}

var (
	sheetName  = "模具库db"
	sheetName1 = "库存-新"
)

func main() {
	fmt.Println("hello world")

	filePath := flag.String("filepath", "文件路径/文件名", "文件路径")
	flag.Parse()

	if len(os.Args) == 1 {
		usage := `Usage:
	-filepath	string		文件路径`
		fmt.Println(usage)
		return
	}

	xlsx, err := excelize.OpenFile(*filePath)
	if err != nil {
		fmt.Println(err)
		return
	}

	var daihaoMap = make(map[string]*mujudaihao)
	var rows [][]string

	for _, v := range xlsx.GetSheetMap() {
		if strings.Contains(v, sheetName) {
			xlsx.DeleteSheet(sheetName)
			continue
		}
		if strings.Compare(v, sheetName1) == 0 {
			rows = xlsx.GetRows(v)
		}
	}

	for i, row := range rows {
		if i == 0 {
			continue
		}
		_, ok := daihaoMap[row[3]]
		if !ok {
			daihaoMap[row[3]] = new(mujudaihao)
		}
		n, err := strconv.Atoi(row[6])
		if err != nil {
			fmt.Println("数量转换错误，行：", i)
			continue
		}
		daihaoMap[row[3]].counts += n
		daihaoMap[row[3]].rows = append(daihaoMap[row[3]].rows, row)
	}

	// for _, v := range daihaoMap {
	// 	fmt.Println(v.counts)
	// 	fmt.Println(v.rows)
	// }
	index := xlsx.NewSheet(sheetName)
	in := 1
	xlsx.SetCellValue(sheetName, "A"+strconv.Itoa(in), "存放仓库")
	xlsx.SetCellValue(sheetName, "B"+strconv.Itoa(in), "保管人")
	xlsx.SetCellValue(sheetName, "C"+strconv.Itoa(in), "状态")

	xlsx.SetCellValue(sheetName, "D"+strconv.Itoa(in), "模具代号")
	xlsx.SetCellValue(sheetName, "E"+strconv.Itoa(in), "模具序号")
	xlsx.SetCellValue(sheetName, "F"+strconv.Itoa(in), "模具名称")
	xlsx.SetCellValue(sheetName, "G"+strconv.Itoa(in), "购入日期")
	xlsx.SetCellValue(sheetName, "H"+strconv.Itoa(in), "启用日期")
	in++

	for k, v := range daihaoMap {
		tmpn := 0
		for _, v1 := range v.rows {
			n1, err := strconv.Atoi(v1[6])
			if err != nil {
				fmt.Println("数量转换错误1：", k, ", ", v1)
				continue
			}
			for i := 0; i < n1; i++ {
				tmpn++
				// xlsx.SetCellValue(index, "A"+strconv.Itoa(in), "存放仓库")
				// xlsx.SetCellValue(index, "B"+strconv.Itoa(in), "保管人")
				// xlsx.SetCellValue(index, "C"+strconv.Itoa(in), "状态")
				xlsx.SetCellValue(sheetName, "D"+strconv.Itoa(in), v1[3])
				xlsx.SetCellValue(sheetName, "E"+strconv.Itoa(in), tmpn)
				xlsx.SetCellValue(sheetName, "F"+strconv.Itoa(in), v1[2])
				xlsx.SetCellValue(sheetName, "G"+strconv.Itoa(in), v1[4])
				xlsx.SetCellValue(sheetName, "H"+strconv.Itoa(in), v1[4])
				in++
			}
		}
		if tmpn != v.counts {
			fmt.Println("数量不对：", k)
		}
	}

	xlsx.SetActiveSheet(index)
	xlsx.SetSheetVisible(sheetName, true)
	xlsx.Save()
}
