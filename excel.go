package go_excel

import (
	"bufio"
	"bytes"
	"errors"
	"fmt"
	"io"
	"reflect"
	"strconv"
	"strings"
	"text/template"
	"time"

	"github.com/xuri/excelize/v2"
)

type Options struct {
	SheetName    string // 表名
	Title        string // 标题
	ShowRemind   bool   // 显示提示
	DefaultStyle bool   // 自定义样式
	SwNum        int64  // 流式写入
}

type Excel struct {
	Fields   map[string]*Column // 字段名称 / 字段
	Option   Options
	RowStyle excelize.Style
	File     *excelize.File
	Sw       *excelize.StreamWriter
	Data     any
}

func New(sheetName, title string) *Excel {
	return &Excel{
		Option: Options{
			SheetName: sheetName,
			Title:     title,
		},
	}
}

func (e *Excel) defaultStyle() {
	titleStyle, _ := e.File.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#DFEBF6"}, Pattern: 1},
		Font: &excelize.Font{
			Bold: true,
			Size: 25,
		},
	})
	e.Sw.SetRow("A1",
		[]any{excelize.Cell{Value: e.Option.Title, StyleID: titleStyle}},
		excelize.RowOpts{Height: 30, Hidden: false})
}

func (e *Excel) export(data any) error {
	columnMap, err := getField(data)
	if err != nil {
		return err
	}
	e.Fields = columnMap
	e.File = excelize.NewFile()
	index, err := e.File.NewSheet(e.Option.SheetName)
	if err != nil {
		return err
	}
	e.Sw, err = e.File.NewStreamWriter(e.Option.SheetName)
	if err != nil {
		return err
	}

	err = e.Sw.SetColWidth(1, 4, 20)
	if err != nil {
		return err
	}
	vCell, _ := numberToLetters(len(e.Fields))
	if err := e.Sw.MergeCell("A1", vCell+"1"); err != nil {
		return err
	}

	e.defaultStyle()

	rvData, ok := e.GetEntityInfo(data)
	if !ok {
		return errors.New("data get entity info err")
	}
	err = e.SetValue(rvData)
	if err != nil {
		return err
	}
	e.Sw.Flush()
	e.File.SetActiveSheet(index)
	_ = e.File.DeleteSheet("Sheet1")
	return nil
}

func (e *Excel) ExportToBytes(data any) ([]byte, error) {
	defer func() {
		if err := e.File.Close(); err != nil {
			return
		}
	}()
	err := e.export(data)
	if err != nil {
		return nil, err
	}
	b := bytes.Buffer{}
	writer := bufio.NewWriter(&b)
	e.File.Write(writer)
	return b.Bytes(), nil
}

func (e *Excel) ExportToFile(data any) error {
	defer func() {
		if err := e.File.Close(); err != nil {
			return
		}
	}()
	err := e.export(data)
	if err != nil {
		return err
	}
	if err := e.File.SaveAs(e.Option.SheetName + ".xlsx"); err != nil {
		return err
	}
	return nil
}

func (e *Excel) GetEntityInfo(data any) (reflect.Value, bool) {
	rv := reflect.ValueOf(data)
	switch rv.Kind() {
	case reflect.Ptr:
		data = rv.Elem().Interface()
		return e.GetEntityInfo(data)
	case reflect.Slice:
		return rv, true
	default:
		return rv, false
	}
}

func (e *Excel) SetValue(rv reflect.Value) error {
	rangeBottoms := "B"
	colList := make([]string, len(e.Fields))
	rangeBottoms, _ = numberToLetters(len(colList))
	rangeBottoms = rangeBottoms + strconv.Itoa(rv.Len()+2)
	for k, v := range e.Fields {
		colList[v.Index] = k
	}
	for i := 2; i < rv.Len()+3; i++ {
		vals := make([]any, 0)
		cell, _ := excelize.CoordinatesToCellName(1, i)
		for _, col := range colList {
			var cellValue any
			if i == 2 {
				cellValue = e.Fields[col].NaturalName
			} else {
				index := i - 3
				rval := reflect.ValueOf(rv.Index(index).Interface())
				if rval.Kind() == reflect.Ptr {
					rval = rval.Elem()
				}
				rfval := rval.FieldByName(e.Fields[col].Field)
				if !rfval.IsZero() {
					x := rfval.Interface()
					cellValue = x
					if e.Fields[col].FieldType == reflect.TypeOf(time.Time{}) {
						cellTime := cellValue.(time.Time)
						cellValue = cellTime.Format("2006-01-02 15:04:05")
					}
				} else {
					cellValue = ""
				}
			}
			vals = append(vals, cellValue)
		}
		err := e.Sw.SetRow(cell, vals)
		if err != nil {
			return err
		}
	}
	err := e.Sw.AddTable(&excelize.Table{
		Range:             "A2:" + rangeBottoms,
		Name:              "excel",
		StyleName:         "TableStyleMedium2",
		ShowFirstColumn:   true,
		ShowLastColumn:    true,
		ShowColumnStripes: true,
	})
	if err != nil {
		return err
	}
	return nil
}

func (e *Excel) report() error {
	defer func() {
		if err := e.File.Close(); err != nil {
			return
		}
	}()

	sheets := e.File.GetSheetList()

	for _, sheet := range sheets {
		rows, err := e.File.GetRows(sheet)
		if err != nil {
			return fmt.Errorf("获取工作表 %s 的行失败: %w", sheet, err)
		}
		rowOffset := 0
		for rowIndex, row := range rows {
			for colIndex, cell := range row {
				if strings.Contains(cell, "{{range") {
					rangeOffset := rowOffset
					rangeValues, err := e.processRangeTemplate(cell)
					if err != nil {
						return fmt.Errorf("处理range模板失败: %w", err)
					}
					for i, value := range rangeValues {
						cellName, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+rangeOffset+i+1)
						if err != nil {
							return fmt.Errorf("转换单元格坐标失败: %w", err)
						}
						if err := e.File.SetCellValue(sheet, cellName, value); err != nil {
							return fmt.Errorf("设置单元格值失败: %w", err)
						}
					}
					rangeOffset += len(rangeValues) - 1
				} else if strings.Contains(cell, "{{") && strings.Contains(cell, "}}") {
					processedValue, err := e.processCellTemplate(cell)
					if err != nil {
						return fmt.Errorf("处理单元格模板失败: %w", err)
					}

					cellName, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+rowOffset+1)
					if err != nil {
						return fmt.Errorf("转换单元格坐标失败: %w", err)
					}
					if err := e.File.SetCellValue(sheet, cellName, processedValue); err != nil {
						return fmt.Errorf("设置单元格值失败: %w", err)
					}
				}
			}
		}
	}
	return nil
}

func (e *Excel) ReportFromFile(filePath string, data any) {
	e.File, _ = excelize.OpenFile(filePath)
	e.Data = data
}

func (e *Excel) ReportFromBytes(reader io.Reader, data any) {
	e.File, _ = excelize.OpenReader(reader)
	e.Data = data
}

func (e *Excel) ReportToFile(outputFile string) error {
	err := e.report()
	if err != nil {
		return err
	}
	if err := e.File.SaveAs(outputFile); err != nil {
		return fmt.Errorf("保存输出文件失败: %w", err)
	}
	return nil
}

func (e *Excel) ReportToBytes() ([]byte, error) {
	err := e.report()
	if err != nil {
		return nil, err
	}
	b := bytes.Buffer{}
	writer := bufio.NewWriter(&b)
	e.File.Write(writer)
	return b.Bytes(), nil
}

func (e *Excel) processCellTemplate(cellContent string) (string, error) {
	tmpl, err := template.New("cell").Parse(cellContent)
	if err != nil {
		return "", fmt.Errorf("解析单元格模板失败: %w", err)
	}

	var buf bytes.Buffer
	if err := tmpl.Execute(&buf, e.Data); err != nil {
		return "", fmt.Errorf("执行单元格模板失败: %w", err)
	}

	return buf.String(), nil
}

func (e *Excel) processRangeTemplate(cellContent string) ([]string, error) {
	tmpl, err := template.New("range").Parse(cellContent)
	if err != nil {
		return nil, fmt.Errorf("解析range模板失败: %w", err)
	}

	var buf bytes.Buffer
	if err := tmpl.Execute(&buf, e.Data); err != nil {
		return nil, fmt.Errorf("执行range模板失败: %w", err)
	}

	// 分割结果，去除空行
	values := strings.Split(buf.String(), "\n")
	var result []string
	for _, v := range values {
		if strings.TrimSpace(v) != "" {
			result = append(result, strings.TrimSpace(v))
		}
	}

	return result, nil
}

func numberToLetters(num int) (string, error) {
	if num <= 0 {
		return "", errors.New("数字必须大于0")
	}

	// 定义字母表
	alphabet := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	base := len(alphabet)

	var result string

	for num > 0 {
		// 计算余数和商
		remainder := (num - 1) % base
		num = (num - 1) / base

		// 将余数对应的字母添加到结果字符串的前面
		result = string(alphabet[remainder]) + result
	}

	return result, nil
}

func refType(value any) (val any, is bool) {
	if value == nil {
		return nil, false
	}
	rv := reflect.ValueOf(value)
	switch rv.Kind() {
	case reflect.Struct:
		return value, true
	case reflect.Ptr:
		if rv.IsNil() {
			return nil, false
		}
		return refType(rv.Elem().Interface())

	case reflect.Slice:
		if rv.Len() > 0 {
			elem := rv.Index(0).Interface()
			return refType(elem)
		}
		return nil, false
	default:
		return nil, false
	}
}
