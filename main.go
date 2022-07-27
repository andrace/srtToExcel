package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"os"
	"strconv"
	"strings"
)

const dirname = `./output`

func main() {
	CreateDir(dirname)
	findDir("./")
}

type srtExcel struct {
	Id       string
	TimeLine string
	Trans1   string
	Trans2   string
	Trans3   string
}

// 遍历的文件夹
func findDir(dir string) {
	fileinfo, err := ioutil.ReadDir(dir)
	if err != nil {
		panic(err)
	}
	// 遍历这个文件夹

	for _, fi := range fileinfo {
		// 判断是不是目录
		var srtList []*srtExcel

		if fi.IsDir() {
			findDir(dir + `/` + fi.Name())
		} else {
			if !strings.Contains(fi.Name(), ".srt") {
				continue
			}
			fmt.Println("读取文件 :", fi.Name())
			file, err := os.Open(fi.Name())
			if err != nil {
				return
			}
			all, err := ioutil.ReadAll(file)
			if err != nil {
				return
			}
			s := string(all)
			split := strings.Split(s, "\r")
			//excel := &srtExcel{}
			lines := 0
			lens := len(split)
			for k, v := range split {
				//fmt.Printf("%T, %v \n", v, v)
				if v == "\n" {
					if k-lines+2 > lens {
						break
					}
					tmp := &srtExcel{}
					tmp.Id = split[k-lines]
					tmp.TimeLine = split[k-lines+1]
					if lines == 3 {
						tmp.Trans1 = split[k-lines+2]
					} else if lines == 4 {
						tmp.Trans1 = split[k-lines+2]
						tmp.Trans2 = split[k-lines+3]
					}
					lines = 0
					srtList = append(srtList, tmp)
				} else {
					lines++
				}
				fmt.Println(lines, v)
			}
			createExcel(fi.Name(), srtList)
			println(`文件名 ：`, fi.Name())

		}
	}
}

func srtToExcel(path string) {

}

func readSrt(file string) (list []*srtExcel) {

	return
}

// HasDir 判断文件夹是否存在
func HasDir(path string) (bool, error) {
	_, _err := os.Stat(path)
	if _err == nil {
		return true, nil
	}
	if os.IsNotExist(_err) {
		return false, nil
	}
	return false, _err
}

// CreateDir 创建文件夹
func CreateDir(path string) {
	_exist, _err := HasDir(path)
	if _err != nil {
		fmt.Printf("获取文件夹异常 -> %v\n", _err)
		return
	}
	if _exist {
		fmt.Println("文件夹已存在！")
	} else {
		err := os.Mkdir(path, os.ModePerm)
		if err != nil {
			fmt.Printf("创建目录异常 -> %v\n", err)
		} else {
			fmt.Println("创建成功!")
		}
	}
}

func createExcel(filename string, list []*srtExcel) {
	fmt.Println("准备写入文件 ："+filename, "一共:", len(list), "条")
	f := excelize.NewFile()
	// 创建一个工作表
	s := "Sheet1"
	index := f.NewSheet(s)
	f.SetActiveSheet(index)
	for k, i := range list {
		// 设置单元格的值
		err := f.SetCellValue(s, "A"+strconv.Itoa(k+1), i.Id)
		if err != nil {
			fmt.Println(err)
		}
		err = f.SetCellValue(s, "B"+strconv.Itoa(k+1), i.TimeLine)
		if err != nil {
			fmt.Println(err)
		}
		err = f.SetCellValue(s, "C"+strconv.Itoa(k+1), i.Trans1)
		if err != nil {
			fmt.Println(err)
		}
		err = f.SetCellValue(s, "D"+strconv.Itoa(k+1), i.Trans2)
		if err != nil {
			fmt.Println(err)
		}
		// 设置工作簿的默认工作表
	}
	// 根据指定路径保存文件
	savePath := dirname + "/" + filename + ".xlsx"
	fmt.Println("save to", savePath)
	if err := f.SaveAs(savePath); err != nil {
		fmt.Println(err)
	}

}
