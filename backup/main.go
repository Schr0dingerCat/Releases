package main

import (
	"flag"
	"fmt"
	"os"
	"sort"
	"strings"
	"time"

	"myBackup/archiver"
	"myBackup/crypto"
	"myBackup/util"
)

var pthSep = string(os.PathSeparator)

//time format "20060102150405.99999"
var formatString = "200601021504"

func main() {
	//	fmt.Println("hello world")
	da := flag.Int("n", 1, "保留最近几个备份")
	sr := flag.String("src", "C:\\Users\\Default", "需要备份的文件或目录")
	ds := flag.String("dst", "C:\\Users\\Default", "需要备份的文件备份保存的目录")

	flag.Parse()
	days := *da
	src := strings.TrimRight(*sr, pthSep)
	dst := strings.TrimRight(*ds, pthSep)

	if len(os.Args) == 1 {
		usage :=
			`Usage:
	-n int
		保留最近几个备份 (default 1)
	-src string
		需要备份的文件或目录 (default "C:\\Users\\Default")
	-dst string
		需要备份的文件备份保存的目录 (default "C:\\Users\\Default")`
		fmt.Println(usage)
		return
	}
	//	fmt.Println(days)
	//	fmt.Println(src)
	//	fmt.Println(dst)

	s := strings.Split(src, pthSep)
	name1 := s[len(s)-1]
	name2 := ".backup.zip"
	//	fmt.Println(name1)

	//获取当前日期
	cdate := time.Now()
	savName1 := cdate.Format(formatString)
	savName := name1 + cdate.Format(formatString) + name2
	//	fmt.Println(savName)
	//备份文件前，已备份最多存在的备份文件名日期后缀
	// delDate := cdate.AddDate(0, 0, 0-*days).Format(formatString)
	existNames := make([]string, days)
	for i := days; i > 0; i-- {
		existNames[i-1] = name1 + cdate.AddDate(0, 0, 0-i).Format(formatString) + name2
	}
	//	fmt.Println(existNames)

	//备份src到dst
	err := archiver.Zip.Make(dst+pthSep+savName+".tmp", []string{src})
	if err != nil {
		fmt.Println("zip err: ", err)
	}
	newHash := crypto.GetFileSha512String(dst + pthSep + savName + ".tmp")

	//存在当前日期备份
	isExist, _ := util.IsExists(dst + pthSep + savName)
	if isExist {
		info, _ := os.Stat(dst + pthSep + savName)
		if !info.IsDir() {
			fmt.Println("文件名为", savName, "的文件已存在")
			//检验sha512是否相同
			oldHash := crypto.GetFileSha512String(dst + pthSep + savName)
			fmt.Println(oldHash)
			fmt.Println(newHash)
			if strings.Compare(newHash, oldHash) == 0 {
				os.Remove(dst + pthSep + savName + ".tmp")
			} else {
				os.Remove(dst + pthSep + savName)
				os.Rename(dst+pthSep+savName+".tmp", dst+pthSep+savName)
			}
		} else {
			del := os.Remove(dst + pthSep + savName)
			fmt.Println("已存在的" + savName + "是目录，已删除")
			if del != nil {
				fmt.Println(del)
			}
		}
	} else {
		os.Rename(dst+pthSep+savName+".tmp", dst+pthSep+savName)
	}
	//遍历保存目录下.zip文件
	fs, _ := util.ListDir(dst, ".zip")
	//	fmt.Println(fs)
	//删除多余备份文件
	var sc []string = make([]string, 0, 10)
	for _, f := range fs {
		f1 := strings.TrimPrefix(f, dst+pthSep)
		if strings.HasPrefix(f1, name1) && strings.HasSuffix(f1, name2) {
			s := strings.TrimSuffix(f1, name2)
			s = strings.TrimPrefix(s, name1)
			_, err := time.ParseInLocation(formatString, s, time.Local)
			if err != nil {
				fmt.Println("不是时间格式: ", s)
				continue
			} else {
				sc = append(sc, s)
			}
		}
	}
	sort.Strings(sc)
	sort.Sort(sort.Reverse(sort.StringSlice(sc)))
	//	fmt.Println(sc)
	var count int = 0
	for i := 0; i < len(sc); i++ {
		info, _ := os.Stat(dst + pthSep + name1 + sc[i] + name2)
		if info.Size() == 152 {
			//			fmt.Println("目录为空")
			del := os.Remove(dst + pthSep + name1 + sc[i] + name2)
			if del != nil {
				fmt.Println("1删除文件错误：", del)
			} else {
				fmt.Println("1已存在的" + name1 + sc[i] + name2 + "，已删除")
			}
			continue
		}
		if strings.Compare(sc[i], savName1) > 0 {
			del := os.Remove(dst + pthSep + name1 + sc[i] + name2)
			if del != nil {
				fmt.Println("2删除文件错误：", del)
			} else {
				fmt.Println("2已存在的" + name1 + sc[i] + name2 + "，已删除")
			}
			continue
		}
		if strings.Compare(sc[i], savName1) == 0 {
			count++
			continue
		}
		if strings.Compare(sc[i], savName1) < 0 {
			h := crypto.GetFileSha512String(dst + pthSep + name1 + sc[i] + name2)
			if strings.Compare(newHash, h) == 0 {
				del := os.Remove(dst + pthSep + name1 + sc[i] + name2)
				if del != nil {
					fmt.Println("3删除文件错误：", del)
				} else {
					fmt.Println("3已存在的" + name1 + sc[i] + name2 + " 的hash相同，已删除")
				}
				continue
			} else {
				//				fmt.Println("hash != ", sc[i])
				if count < days {
					count++
					continue
				} else {
					del := os.Remove(dst + pthSep + name1 + sc[i] + name2)
					if del != nil {
						fmt.Println("4删除文件错误：", del)
					} else {
						fmt.Println("4已存在的" + name1 + sc[i] + name2 + "，已删除")
					}
				}
			}
		}
	}
}
