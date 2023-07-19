package main

import (
	"bytes"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path"
	"path/filepath"
	"strconv"
	"strings"
)

func main() {
	var args = os.Args

	/**
	获取入参,然后进行切换,第一个参数是excel,第二个为大小,第三个为excel文件夹的切分后名称前缀,默认和原始一致,第四个为sheet名称,默认为Sheet1
	*/
	argCount := len(args)
	if argCount == 1 {
		println("excel abs path must be input,like:\n ./excel excelPath pageSize newExcelPrefix sheetName,\n eg: ./excel /tmp/a.xlsx 2 c Sheet1")
		return
	}
	// 1.获取文件
	var excelAbs = args[1]
	f, err := excelize.OpenFile(excelAbs)
	if err != nil {
		println(err.Error())
		return
	}

	//2.获取切分大小
	var pageSize = "5000"
	if argCount >= 3 {
		pageSize = args[2]
	}

	num, err := strconv.Atoi(pageSize)
	//如果有名称,那么按照名称,没有,那么按照表名
	if err != nil {
		println(err.Error())
		return
	}

	//3.获取切分文件前缀
	var ext = path.Ext(excelAbs)
	var base = path.Base(excelAbs)
	base, _, _ = strings.Cut(excelAbs, ext)

	//如果没有,那么创建一个目录,文件夹为archive
	distDir := base + "_archive"
	err = os.MkdirAll(distDir, os.ModePerm)
	if err != nil {
		fmt.Println(err)
		return
	}

	namePrefix := filepath.Join(distDir, base)
	if argCount >= 4 {
		var a3 = os.Args[3]
		if a3 != "" {
			namePrefix = filepath.Join(distDir, a3)
		}
	}

	//4.获取sheet名称
	var sheetName = "Sheet1"
	if argCount >= 5 {
		sheetName = args[4]
	}
	///opt/work/life/我们的情薄.xlsx 2 a Sheet1 2
	//5.表头个数
	var headerCount = "1"
	if argCount >= 6 {
		headerCount = args[5]
	}
	templateHeaderCount, err := strconv.Atoi(headerCount)
	//如果有名称,那么按照名称,没有,那么按照表名
	if err != nil {
		println(err.Error())
		return
	}

	// 获取 Sheet1 上所有单元格记录
	rows, err := f.GetRows(sheetName)
	if err != nil {
		println(err.Error())
		return
	}

	var total = len(rows) - templateHeaderCount
	var count = int(total / num)
	var left = int(total % num)
	//如果大于0,那么就是还需要再加一个表格
	if left > 0 {
		left = 1
	}

	fmt.Printf("excelAbs:%s,records:%d,pageSize:%d,table count:%d\n", excelAbs, total, num, count+left)
	/**
	获取第一行的表头,然后将这些改成不同的文件,写入第一行,
	*/
	var index int = 0
	colums := make([][]string, templateHeaderCount)

	var file *excelize.File
	defer func() {
		if err := file.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	var record int = templateHeaderCount
	for idx, row := range rows {
		//创建文件名称
		if idx < (templateHeaderCount - 1) {
			colums[idx] = row
			continue
		}
		if idx == (templateHeaderCount - 1) {
			colums[idx] = row
			file = newFile(colums)
			continue
		}
		cell, err := excelize.CoordinatesToCellName(1, record+1)
		if err != nil {
			fmt.Println(err)
			return
		}
		file.SetSheetRow("Sheet1", cell, &row)
		record = record + 1
		if idx%num == 0 {
			index = index + 1
			save(namePrefix, index, file)
			file = newFile(colums)
			record = templateHeaderCount
			fmt.Printf("pageSize:%d,process count:%d\n", num, idx+1)
		} else if idx == total {
			index = index + 1
			save(namePrefix, index, file)
			fmt.Printf("process finish count:%d\n", idx)
			return
		}
	}
}

func newFile(colums [][]string) *excelize.File {
	file := excelize.NewFile()
	for a, b := range colums {
		for i, v := range b {
			cell, err := excelize.CoordinatesToCellName(i+1, a+1)
			if err != nil {
				fmt.Println(err)
				return nil
			}
			file.SetCellValue("Sheet1", cell, v)
		}
	}
	return file
}

func save(namePrefix string, index int, file *excelize.File) {

	//当最后一次,只做保存
	var buf bytes.Buffer
	buf.WriteString(namePrefix)
	buf.WriteString("_")
	buf.WriteString(strconv.Itoa(index))
	buf.WriteString(".xlsx")
	if err := file.SaveAs(buf.String()); err != nil {
		fmt.Println(err)
	}
}
