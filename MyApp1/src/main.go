package main

import (
	"flag"
	"fmt"
	"os"
	"strconv"
	"strings"

	"excelize"
)

var (
	index2016 = "2016帐龄"
	index2017 = "2017帐龄"
	index2018 = "2018帐龄"
	i2016     = 1
	i2017     = 1
	i2018     = 1
)

type huopin struct {
	name       string
	jiezhuan15 []string
	jiezhuan16 []string
	jiezhuan17 []string
	isruku16   bool
	isruku17   bool
	isruku18   bool
	jiecun16   [][]string
	jiecun17   [][]string
	jiecun18   [][]string
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
	-filePath string	wenjianjueduilujjing`
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
		if strings.Compare(v, index2016) == 0 || strings.Compare(v, index2017) == 0 || strings.Compare(v, index2018) == 0 || strings.Compare(v, "Other") == 0 {
			xlsx.DeleteSheet(v)
			continue
			//			xlsx.Save()
		}
		//		fmt.Println("leneeee:    ", len(xlsx.GetRows(v)))
		for i, row := range xlsx.GetRows(v) {
			if i == 0 {
				continue
			}
			// 末尾有的有空格
			row[0] = strings.TrimSpace(row[0])
			_, ok := huopinMap[row[0]]
			if !ok {
				huopinMap[row[0]] = new(huopin)
				huopinMap[row[0]].name = row[0]
			}
			if strings.Contains(row[2], "2015年结转") {
				huopinMap[row[0]].jiezhuan15 = row
			}
			if strings.Contains(row[2], "2016年结转") {
				huopinMap[row[0]].jiezhuan16 = row
			}
			if strings.Contains(row[2], "2017年结转") {
				huopinMap[row[0]].jiezhuan17 = row
			}
			if strings.Contains(row[2], "月_结存") {

				if strings.Compare(row[1], "2016") == 0 {
					huopinMap[row[0]].jiecun16 = append(huopinMap[row[0]].jiecun16, row)
				}
				if strings.Compare(row[1], "2017") == 0 {
					huopinMap[row[0]].jiecun17 = append(huopinMap[row[0]].jiecun17, row)
				}
				if strings.Compare(row[1], "2018") == 0 {
					huopinMap[row[0]].jiecun18 = append(huopinMap[row[0]].jiecun18, row)
				}
			}
			if strings.Compare(row[3], "0") != 0 {
				if strings.Contains(row[1], "2016") {
					huopinMap[row[0]].isruku16 = true
				}
				if strings.Contains(row[1], "2017") {
					huopinMap[row[0]].isruku17 = true
				}
				if strings.Contains(row[1], "2018") {
					huopinMap[row[0]].isruku18 = true
				}
			}
		}
	}
	//chuligegehuopin
	index := xlsx.NewSheet(index2016)
	xlsx.NewSheet(index2017)
	xlsx.NewSheet(index2018)
	xlsx.SetCellValue(index2016, "A"+strconv.Itoa(i2016), "场内编号")
	xlsx.SetCellValue(index2016, "B"+strconv.Itoa(i2016), "年份")
	xlsx.SetCellValue(index2016, "C"+strconv.Itoa(i2016), "摘要")
	xlsx.SetCellValue(index2016, "D"+strconv.Itoa(i2016), "入库数量")
	xlsx.SetCellValue(index2016, "E"+strconv.Itoa(i2016), "库存数量")
	xlsx.SetCellValue(index2016, "F"+strconv.Itoa(i2016), "帐龄")
	i2016++
	xlsx.SetCellValue(index2017, "A"+strconv.Itoa(i2017), "场内编号")
	xlsx.SetCellValue(index2017, "B"+strconv.Itoa(i2017), "年份")
	xlsx.SetCellValue(index2017, "C"+strconv.Itoa(i2017), "摘要")
	xlsx.SetCellValue(index2017, "D"+strconv.Itoa(i2017), "入库数量")
	xlsx.SetCellValue(index2017, "E"+strconv.Itoa(i2017), "库存数量")
	xlsx.SetCellValue(index2017, "F"+strconv.Itoa(i2017), "帐龄")
	i2017++
	xlsx.SetCellValue(index2018, "A"+strconv.Itoa(i2018), "场内编号")
	xlsx.SetCellValue(index2018, "B"+strconv.Itoa(i2018), "年份")
	xlsx.SetCellValue(index2018, "C"+strconv.Itoa(i2018), "摘要")
	xlsx.SetCellValue(index2018, "D"+strconv.Itoa(i2018), "入库数量")
	xlsx.SetCellValue(index2018, "E"+strconv.Itoa(i2018), "库存数量")
	xlsx.SetCellValue(index2018, "F"+strconv.Itoa(i2018), "帐龄")
	i2018++

	q := 1
	xlsx.NewSheet("Other")
	xlsx.SetCellValue("Other", "A"+strconv.Itoa(q), "标记")
	xlsx.SetCellValue("Other", "B"+strconv.Itoa(q), "厂内编号")
	q++

	for k, v := range huopinMap {
		if len(v.jiecun16) < 1 && len(v.jiecun17) < 1 && len(v.jiecun18) < 1 {
			xlsx.SetCellValue("Other", "A"+strconv.Itoa(q), "3年没结存")
			xlsx.SetCellValue("Other", "B"+strconv.Itoa(q), v.name)
			xlsx.SetCellValue("Other", "C"+strconv.Itoa(q), len(v.jiecun16))
			xlsx.SetCellValue("Other", "D"+strconv.Itoa(q), len(v.jiecun17))
			xlsx.SetCellValue("Other", "E"+strconv.Itoa(q), len(v.jiecun18))
			xlsx.SetCellValue("Other", "F"+strconv.Itoa(q), v.isruku16)
			xlsx.SetCellValue("Other", "G"+strconv.Itoa(q), v.isruku17)
			xlsx.SetCellValue("Other", "H"+strconv.Itoa(q), v.isruku18)
			q++
			continue
		}
		//2016年帐龄
		if len(v.jiecun16) > 0 {
			MyJiangXu(v.jiecun16, 2)
			if len(v.jiezhuan16) > 0 {
				if strings.Compare(v.jiecun16[0][4], v.jiezhuan16[4]) != 0 {
					xlsx.SetCellValue("Other", "A"+strconv.Itoa(q), "2016最后结存不等于2016结转")
					xlsx.SetCellValue("Other", "B"+strconv.Itoa(q), v.name)
					q++
				}
			}
			//最后一个月结存
			if strings.Compare(v.jiecun16[0][4], "0") != 0 {
				xlsx.SetCellValue(index2016, "A"+strconv.Itoa(i2016), v.jiecun16[0][0])
				xlsx.SetCellValue(index2016, "B"+strconv.Itoa(i2016), v.jiecun16[0][1])
				xlsx.SetCellValue(index2016, "C"+strconv.Itoa(i2016), v.jiecun16[0][2])
				xlsx.SetCellValue(index2016, "D"+strconv.Itoa(i2016), v.jiecun16[0][3])
				xlsx.SetCellValue(index2016, "E"+strconv.Itoa(i2016), v.jiecun16[0][4])
				if v.isruku16 {
					xlsx.SetCellValue(index2016, "F"+strconv.Itoa(i2016), "1年")
				} else {
					xlsx.SetCellValue(index2016, "F"+strconv.Itoa(i2016), "1年以上")
				}
				i2016++
			}
		}
		//2017帐龄
		if len(v.jiecun17) > 0 {
			MyJiangXu(v.jiecun17, 2)
			if len(v.jiezhuan17) > 0 {
				if strings.Compare(v.jiecun17[0][4], v.jiezhuan17[4]) != 0 {
					xlsx.SetCellValue("Other", "A"+strconv.Itoa(q), "2017最后结存不等于2017结转")
					xlsx.SetCellValue("Other", "B"+strconv.Itoa(q), v.name)
					q++
				}
			}
			//最后一个月结存
			if strings.Compare(v.jiecun17[0][4], "0") != 0 {
				xlsx.SetCellValue(index2017, "A"+strconv.Itoa(i2017), v.jiecun17[0][0])
				xlsx.SetCellValue(index2017, "B"+strconv.Itoa(i2017), v.jiecun17[0][1])
				xlsx.SetCellValue(index2017, "C"+strconv.Itoa(i2017), v.jiecun17[0][2])
				xlsx.SetCellValue(index2017, "D"+strconv.Itoa(i2017), v.jiecun17[0][3])
				xlsx.SetCellValue(index2017, "E"+strconv.Itoa(i2017), v.jiecun17[0][4])
				if v.isruku17 {
					xlsx.SetCellValue(index2017, "F"+strconv.Itoa(i2017), "1年")
				} else {
					if v.isruku16 {
						xlsx.SetCellValue(index2017, "F"+strconv.Itoa(i2017), "2年")
					} else {
						xlsx.SetCellValue(index2017, "F"+strconv.Itoa(i2017), "2年以上")
					}
				}
				i2017++
			}
		}
		//2018帐龄
		if len(v.jiecun18) > 0 {
			MyJiangXu(v.jiecun18, 2)
			//最后一个月结存
			if strings.Compare(v.jiecun18[0][4], "0") != 0 {
				xlsx.SetCellValue(index2018, "A"+strconv.Itoa(i2018), v.jiecun18[0][0])
				xlsx.SetCellValue(index2018, "B"+strconv.Itoa(i2018), v.jiecun18[0][1])
				xlsx.SetCellValue(index2018, "C"+strconv.Itoa(i2018), v.jiecun18[0][2])
				xlsx.SetCellValue(index2018, "D"+strconv.Itoa(i2018), v.jiecun18[0][3])
				xlsx.SetCellValue(index2018, "E"+strconv.Itoa(i2018), v.jiecun18[0][4])
				if v.isruku18 {
					xlsx.SetCellValue(index2018, "F"+strconv.Itoa(i2018), "1年")
				} else {
					if v.isruku17 {
						xlsx.SetCellValue(index2018, "F"+strconv.Itoa(i2018), "2年")
					} else {
						if v.isruku16 {
							xlsx.SetCellValue(index2018, "F"+strconv.Itoa(i2018), "3年")
						} else {
							xlsx.SetCellValue(index2018, "F"+strconv.Itoa(i2018), "3年以上")
						}
					}
				}
				i2018++
			}
		}
	}
	xlsx.SetActiveSheet(index)
	xlsx.SetSheetVisible(index2016, true)
	xlsx.Save()
}
