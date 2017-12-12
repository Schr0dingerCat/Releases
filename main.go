package main

import (
	"fmt"
	"os"
	"strconv"

	"github.com/excelize"
)

func main() {
	fmt.Println("hello world")
	if len(os.Args) != 2 {
		fmt.Printf("Usage:\n\t%s excelFileName\n", os.Args[0])
		os.Exit(1)
	}

	fileName := os.Args[1]
	fmt.Println(fileName)
	xlsx, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	result := "Result"
	var list = make(map[string][]string)
	for _, v := range xlsx.GetSheetMap() {
		if v == result {
			xlsx.DeleteSheet(v)
			xlsx.Save()
			continue
		}
		for _, row := range xlsx.GetRows(v) {
			if row[0] != "" {
				list[row[0]+row[1]+row[3]+row[2]] = row[0:4]
			}
		}
	}

	index := xlsx.NewSheet(result)
	i := 1
	for _, v := range list {
		xlsx.SetCellValue(result, "A"+strconv.Itoa(i), v[0])
		xlsx.SetCellValue(result, "B"+strconv.Itoa(i), v[1])
		xlsx.SetCellValue(result, "C"+strconv.Itoa(i), v[2])
		xlsx.SetCellValue(result, "D"+strconv.Itoa(i), v[3])
		i++
	}
	xlsx.SetActiveSheet(index)
	xlsx.SetSheetVisible(result, true)
	xlsx.Save()
	fmt.Println("ok")
}
