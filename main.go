/*
MIT License

# Copyright (c) 2023 WittonBell

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/
package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/xuri/excelize/v2"
)

var taskIDField string
var taskTitleField string
var preTaskField string
var noCaseSensitive bool // 是否不区分大小写
var fieldNameRow uint    // 字段名所在行号
var dataStartRow uint    // 数据开始行号

type node struct {
	taskID    string
	taskTitle string
	preTaskID string
}

type multiMap map[string][]*node

func (slf multiMap) Add(key string, nd *node) {
	if len(slf) == 0 {
		slf[key] = []*node{nd}
	} else {
		slf[key] = append(slf[key], nd)
	}
}

func (slf multiMap) Get(key string) []*node {
	if slf == nil {
		return nil
	}
	return slf[key]
}

func (slf multiMap) Del(key string) {
	delete(slf, key)
}

func searchKeyCol(rows *excelize.Rows) (TaskIDCol, PreTaskIDCol, TitleCol int) {
	row, err := rows.Columns()
	if err != nil {
		fmt.Println(err.Error())
	}

	for i, col := range row {
		name := col
		if noCaseSensitive {
			name = strings.ToLower(col)
		}
		if name == preTaskField {
			PreTaskIDCol = i + 1
		} else if name == taskIDField {
			TaskIDCol = i + 1
		} else if name == taskTitleField {
			TitleCol = i + 1
		}
	}
	return
}

func readExcel(filePath string) (map[string]*node, multiMap) {
	fd, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Printf("读取文件`%s`失败", filePath)
		return nil, nil
	}
	defer func() {
		fd.Close()
	}()
	TaskIDCol, PreTaskIDCol, TitleCol := -1, -1, -1
	sheetName := fd.GetSheetName(0)
	rows, err := fd.Rows(sheetName)
	if err != nil {
		return nil, nil
	}
	defer func() {
		rows.Close()
	}()

	m := map[string]*node{}
	mm := multiMap{}

	for i := 1; rows.Next(); i++ {
		if i == int(fieldNameRow) {
			TaskIDCol, PreTaskIDCol, TitleCol = searchKeyCol(rows)
			isOk := true
			if TaskIDCol < 0 {
				isOk = false
				fmt.Printf("要求字段名：%s\n", taskIDField)
			}
			if PreTaskIDCol < 0 {
				isOk = false
				fmt.Printf("要求字段名：%s\n", preTaskField)
			}
			if TitleCol < 0 {
				isOk = false
				fmt.Printf("要求字段名：%s\n", taskTitleField)
			}
			if !isOk {
				return nil, nil
			}
		}
		if i < int(dataStartRow) {
			continue
		}
		TaskIDCell, err := excelize.CoordinatesToCellName(TaskIDCol, i)
		if err != nil {
			continue
		}
		PreTaskIDCell, err := excelize.CoordinatesToCellName(PreTaskIDCol, i)
		if err != nil {
			continue
		}

		TitleColCell, err := excelize.CoordinatesToCellName(TitleCol, i)
		if err != nil {
			continue
		}

		TaskID, err := fd.GetCellValue(sheetName, TaskIDCell)
		if err != nil || TaskID == "" {
			continue
		}

		Title, err := fd.GetCellValue(sheetName, TitleColCell)
		if err != nil || Title == "" {
			continue
		}

		PreTaskID, err := fd.GetCellValue(sheetName, PreTaskIDCell)
		if err != nil {
			continue
		}

		if PreTaskID == "" {
			PreTaskID = "0"
		}

		nd := &node{taskID: TaskID, taskTitle: Title, preTaskID: PreTaskID}
		mm.Add(PreTaskID, nd)
		m[TaskID] = nd
	}

	return m, mm
}

func usage() {
	w := flag.CommandLine.Output()
	fmt.Fprintf(w, "%s 应用程序是将Excel任务表中的关系转换成Markdown的mermaid图，方便使用Markdown工具直观地查看任务依赖。", filepath.Base(os.Args[0]))
	fmt.Fprintln(w)
	fmt.Fprintf(w, "命令格式：%s -hr [字段所在行号] -dr [数据起始行号] [-nc] -id [任务ID字段名] -t [任务标题字段名] -pid [前置任务ID字段名] -o <输出文件> <Excel文件路径>", filepath.Base(os.Args[0]))
	fmt.Fprintln(w)
	flag.CommandLine.PrintDefaults()
	fmt.Fprintln(w, "  -h")
	fmt.Fprintln(w, "    \t显示此帮助")
}

func main() {
	var outputFile string

	flag.CommandLine.Usage = usage
	flag.BoolVar(&noCaseSensitive, "nc", false, "字段名不区分大小写")
	flag.UintVar(&fieldNameRow, "hr", 1, "字段所在行号")
	flag.UintVar(&dataStartRow, "dr", 2, "数据起始行号")
	flag.StringVar(&taskIDField, "id", "ID", "-id [任务ID字段名]")
	flag.StringVar(&taskTitleField, "t", "Title", "-t [任务标题字段名]")
	flag.StringVar(&preTaskField, "pid", "PreTask", "-pid [前置任务ID字段名]")
	flag.StringVar(&outputFile, "o", "任务图.md", "-o <输出文件>")

	flag.Parse()
	if flag.NArg() < 1 {
		usage()
		return
	}
	if noCaseSensitive {
		taskIDField = strings.ToLower(taskIDField)
		taskTitleField = strings.ToLower(taskTitleField)
		preTaskField = strings.ToLower(preTaskField)
	}
	m, mm := readExcel(flag.Arg(0))
	buildGraph(m, mm, outputFile)
}

type nodeStack struct {
	ar []*node
}

func (slf nodeStack) Len() int {
	return len(slf.ar)
}

func (slf nodeStack) Less(i, j int) bool {
	// 由于在处理时是从栈顶开始的，栈是先进后出，所以要想按从小到大的顺序处理，这里就得按从大到小排序
	// 即大的先压栈，小的后压栈。
	if slf.ar[i].preTaskID == slf.ar[j].preTaskID {
		return slf.ar[i].taskID > slf.ar[j].taskID
	}
	return slf.ar[i].preTaskID > slf.ar[j].preTaskID
}

func (slf nodeStack) Swap(i, j int) {
	slf.ar[i], slf.ar[j] = slf.ar[j], slf.ar[i]
}

func (slf *nodeStack) Push(nd *node) {
	slf.ar = append(slf.ar, nd)
}

func (slf *nodeStack) Pop() *node {
	size := len(slf.ar)
	if size == 0 {
		return nil
	}
	nd := slf.ar[size-1]
	slf.ar = slf.ar[:size-1]
	return nd
}

func (slf *nodeStack) Concat(s nodeStack) {
	slf.ar = append(slf.ar, s.ar...)
}

func buildGraph(m map[string]*node, mm multiMap, outputFile string) {
	graph := "```mermaid\ngraph TB\n"
	graph += "subgraph  \n"
	var stack nodeStack
	ar := mm.Get("0")
	for _, v := range ar {
		stack.Push(v)
	}
	// 从大到小排序
	sort.Sort(stack)
	//添加根节点
	m["0"] = &node{taskTitle: "无效任务ID"}

	topSize := stack.Len()
	// 保存各层的节点数
	arSize := []int{topSize}
	layer := 0
	hasData := topSize > 0
	topSize--

	for {
		v := stack.Pop()
		if v == nil {
			if hasData {
				graph += "end\n"
				fmt.Printf("end\n")
			}
			break
		}
		arSize[layer]--
		if layer == 0 {
			// 如果顶层节点的数量发生变化，则代表subgraph结束了
			if arSize[layer] < topSize {
				graph += "end\n"
				fmt.Printf("end\n")
			}
		}

		x := m[v.preTaskID]
		graph += fmt.Sprintf("%s:%s --> %s:%s\n", v.preTaskID, x.taskTitle, v.taskID, v.taskTitle)
		fmt.Printf("%s:%s --> %s:%s\n", v.preTaskID, x.taskTitle, v.taskID, v.taskTitle)

		if layer == 0 {
			graph += "subgraph  \n"
			fmt.Printf("subgraph  \n")
		}
		if layer > 0 && arSize[layer] == 0 {
			arSize = arSize[:layer]
			layer--
			topSize = arSize[0]
		}

		ar = mm.Get(v.taskID)
		if ar == nil {
			// 没有任务以v.taskID为前置任务的了
			continue
		}
		// 将下级任务压入新的栈，进行排序，以保证所有下级任务都是按从小到大的顺序排列
		var st nodeStack
		for _, v := range ar {
			st.Push(v)
		}
		sort.Sort(st)
		stack.Concat(st)

		// 保存该层的节点数
		arSize = append(arSize, st.Len())
		layer++
	}
	graph += "end\n"
	graph += "```"

	os.WriteFile(outputFile, []byte(graph), os.ModePerm)
	fmt.Println("完成")
}
