package main

import (
	"flag"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/excelize"
)

var (
	result = "Result"
)

type huopin struct {
	name      string
	shangnian []string
	jiecun    [][]string
	ruku      [][]string
}

func MyJiangXu(src [][]string, n int) {
	if len(src) == 1 {
		return
	}
	for i := 0; i < len(src)-1; i++ {
		for j := i + 1; j < len(src); j++ {
			if src[i][n] < src[j][n] {
				src[i], src[j] = src[j], src[i]
			}
		}
	}
}

func main() {
	f1 := flag.String("filePath", "/home/abc/workspace/2016长和1号仓库年末报表2016年.xlsx", "wen jian pu jing")
	flag.Parse()

	if len(os.Args) == 1 {
		usage := `Usage:
				-filePath string
						wenjianjueduilujjing`
		fmt.Println(usage)
		return
	}

	xlsx, err := excelize.OpenFile(*f1)
	//	xlsx, err := excelize.OpenFile("/home/abc/workspace/2016长和1号仓库年末报表2016年.xlsx")
	if err != nil {
		fmt.Println(err)
	}

	var huopinMap = make(map[string]*huopin)

	for _, v := range xlsx.GetSheetMap() {
		if strings.Compare(v, result) == 0 || strings.Compare(v, "Other") == 0 {
			xlsx.DeleteSheet(v)
			//			xlsx.Save()
		}
		//		fmt.Println("leneeee:    ", len(xlsx.GetRows(v)))
		for i, row := range xlsx.GetRows(v) {
			if i == 0 {
				continue
			}
			_, ok := huopinMap[row[0]]
			if !ok {
				huopinMap[row[0]] = new(huopin)
				huopinMap[row[0]].name = row[0]
			}
			if strings.Contains(row[3], "_上年") {
				huopinMap[row[0]].shangnian = row
			}
			if strings.Contains(row[3], "月_结存") {
				huopinMap[row[0]].jiecun = append(huopinMap[row[0]].jiecun, row)
			}
			if strings.Compare(row[4], "0") != 0 {
				huopinMap[row[0]].ruku = append(huopinMap[row[0]].ruku, row)
			}
		}
	}
	//chuligegehuopin
	r := 1
	q := 1
	index := xlsx.NewSheet(result)
	xlsx.NewSheet("Other")
	xlsx.SetCellValue(result, "A"+strconv.Itoa(r), "场内编号")
	xlsx.SetCellValue(result, "B"+strconv.Itoa(r), "编号")
	xlsx.SetCellValue(result, "C"+strconv.Itoa(r), "日期")
	xlsx.SetCellValue(result, "D"+strconv.Itoa(r), "摘要")
	xlsx.SetCellValue(result, "E"+strconv.Itoa(r), "入库数量")
	xlsx.SetCellValue(result, "F"+strconv.Itoa(r), "出库数量")
	xlsx.SetCellValue(result, "G"+strconv.Itoa(r), "库存数量")
	xlsx.SetCellValue(result, "H"+strconv.Itoa(r), "备注")
	xlsx.SetCellValue(result, "I"+strconv.Itoa(r), "标记")
	r++
	xlsx.SetCellValue("Other", "A"+strconv.Itoa(q), "标记")
	xlsx.SetCellValue("Other", "B"+strconv.Itoa(q), "场内编号")
	q++
	for _, v := range huopinMap {
		//		fmt.Println(v)
		if len(v.jiecun) < 1 {
			fmt.Println("没有结存：", v.name)
			xlsx.SetCellValue("Other", "A"+strconv.Itoa(q), "没有结存")
			xlsx.SetCellValue("Other", "B"+strconv.Itoa(q), v.name)
			q++
			continue
		}
		MyJiangXu(v.jiecun, 3)
		if len(v.jiecun[0][6]) == 0 {
			fmt.Println("最后一个月结存为空：", v.name)
			xlsx.SetCellValue("Other", "A"+strconv.Itoa(q), "最后一个月结存为空")
			xlsx.SetCellValue("Other", "B"+strconv.Itoa(q), v.name)
			q++
			continue
		} else if strings.Compare(v.jiecun[0][6], "0") == 0 {
			fmt.Println("最后一个月结存为0：", v.name)
			xlsx.SetCellValue("Other", "A"+strconv.Itoa(q), "最后一个月结存为0")
			xlsx.SetCellValue("Other", "B"+strconv.Itoa(q), v.name)
			q++
			continue
		} else if strings.Contains(v.jiecun[0][6], "-") {
			fmt.Println("最后一个月结存为负：", v.name)
			xlsx.SetCellValue("Other", "A"+strconv.Itoa(q), "最后一个月结存为负")
			xlsx.SetCellValue("Other", "B"+strconv.Itoa(q), v.name)
			q++
			continue
		}
		num, err := strconv.Atoi(v.jiecun[0][6])
		if err != nil {
			fmt.Println("结存数转换失败", v.name)
			xlsx.SetCellValue("Other", "A"+strconv.Itoa(q), "结存数转换失败")
			xlsx.SetCellValue("Other", "B"+strconv.Itoa(q), v.name)
			q++
			continue
		}
		//最后月结存不为零
		xlsx.SetCellValue(result, "A"+strconv.Itoa(r), v.jiecun[0][0])
		xlsx.SetCellValue(result, "B"+strconv.Itoa(r), v.jiecun[0][1])
		xlsx.SetCellValue(result, "C"+strconv.Itoa(r), v.jiecun[0][2])
		xlsx.SetCellValue(result, "D"+strconv.Itoa(r), v.jiecun[0][3])
		xlsx.SetCellValue(result, "E"+strconv.Itoa(r), v.jiecun[0][4])
		xlsx.SetCellValue(result, "F"+strconv.Itoa(r), v.jiecun[0][5])
		xlsx.SetCellValue(result, "G"+strconv.Itoa(r), v.jiecun[0][6])
		xlsx.SetCellValue(result, "H"+strconv.Itoa(r), v.jiecun[0][7])
		if len(v.ruku) == 0 {
			xlsx.SetCellValue(result, "I"+strconv.Itoa(r), "本年没有入库")
		}
		r++
		if len(v.ruku) > 0 {
			MyJiangXu(v.ruku, 2)
		} else {
			continue
		}
		if len(v.shangnian) == 0 {
			fmt.Println("上年结存数据没有：", v.name)
			xlsx.SetCellValue("Other", "A"+strconv.Itoa(q), "上年结存数据没有")
			xlsx.SetCellValue("Other", "B"+strconv.Itoa(q), v.name)
			q++
		} else {
			xlsx.SetCellValue(result, "A"+strconv.Itoa(r), v.shangnian[0])
			xlsx.SetCellValue(result, "B"+strconv.Itoa(r), v.shangnian[1])
			xlsx.SetCellValue(result, "C"+strconv.Itoa(r), v.shangnian[2])
			xlsx.SetCellValue(result, "D"+strconv.Itoa(r), v.shangnian[3])
			xlsx.SetCellValue(result, "E"+strconv.Itoa(r), v.shangnian[4])
			xlsx.SetCellValue(result, "F"+strconv.Itoa(r), v.shangnian[5])
			xlsx.SetCellValue(result, "G"+strconv.Itoa(r), v.shangnian[6])
			xlsx.SetCellValue(result, "H"+strconv.Itoa(r), v.shangnian[7])
			r++
		}
		n := 0
		for i := 0; i < len(v.ruku); i++ {
			m, err := strconv.Atoi(v.ruku[i][4])
			if err != nil {
				fmt.Println("入库数转换失败", v.name)
				xlsx.SetCellValue("Other", "A"+strconv.Itoa(q), "入库数转换失败")
				xlsx.SetCellValue("Other", "B"+strconv.Itoa(q), v.name)
				q++
				continue
			}
			n += m
			xlsx.SetCellValue(result, "A"+strconv.Itoa(r), v.ruku[i][0])
			xlsx.SetCellValue(result, "B"+strconv.Itoa(r), v.ruku[i][1])
			xlsx.SetCellValue(result, "C"+strconv.Itoa(r), v.ruku[i][2])
			xlsx.SetCellValue(result, "D"+strconv.Itoa(r), v.ruku[i][3])
			xlsx.SetCellValue(result, "E"+strconv.Itoa(r), v.ruku[i][4])
			xlsx.SetCellValue(result, "F"+strconv.Itoa(r), v.ruku[i][5])
			xlsx.SetCellValue(result, "G"+strconv.Itoa(r), v.ruku[i][6])
			xlsx.SetCellValue(result, "H"+strconv.Itoa(r), v.ruku[i][7])
			r++
			if n >= num {
				break
			}
		}
		if n < num {
			fmt.Println("本年入库数小于年底结存数：", v.name)
			xlsx.SetCellValue("Other", "A"+strconv.Itoa(q), "本年入库数小于年底结存数")
			xlsx.SetCellValue("Other", "B"+strconv.Itoa(q), v.name)
			q++
		}
	}
	xlsx.SetActiveSheet(index)
	xlsx.SetSheetVisible(result, true)
	xlsx.Save()
}
