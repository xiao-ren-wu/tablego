package tablego

import (
	"encoding/csv"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"

	"github.com/xuri/excelize/v2"
)

type BaseWriter struct {
	filePath string
	headers  []string
	data     [][]string
}

type XlsxWriter struct {
	*BaseWriter
	file *excelize.File
}

type CsvWriter struct {
	*BaseWriter
	file *os.File
}

// NewTableWriter 根据文件后缀创建对应的writer
func NewTableWriter(filePath string) (TableWriter, error) {
	ext := strings.ToLower(filepath.Ext(filePath))
	switch ext {
	case ".xlsx":
		return newXlsxWriter(filePath)
	case ".csv":
		return newCsvWriter(filePath)
	default:
		return nil, fmt.Errorf("unsupported file type: %s", ext)
	}
}

func newXlsxWriter(filePath string) (*XlsxWriter, error) {
	f := excelize.NewFile()
	return &XlsxWriter{
		BaseWriter: &BaseWriter{
			filePath: filePath,
			data:     make([][]string, 0),
		},
		file: f,
	}, nil
}

func newCsvWriter(filePath string) (*CsvWriter, error) {
	f, err := os.Create(filePath)
	if err != nil {
		return nil, err
	}
	return &CsvWriter{
		BaseWriter: &BaseWriter{
			filePath: filePath,
			data:     make([][]string, 0),
		},
		file: f,
	}, nil
}

// WriteLine 写入一行数据
func (w *BaseWriter) WriteLine(record any) error {
	row, err := w.parseRecord(record)
	if err != nil {
		return err
	}
	w.data = append(w.data, row)
	return nil
}

// parseRecord 解析记录为字符串切片
func (w *BaseWriter) parseRecord(record any) ([]string, error) {
	switch v := record.(type) {
	case map[string]any:
		return w.parseMap(v)
	default:
		return w.parseStruct(record)
	}
}

// parseMap 解析map类型的记录
func (w *BaseWriter) parseMap(m map[string]any) ([]string, error) {
	if len(w.headers) == 0 {
		// 第一行数据，设置表头
		w.headers = make([]string, 0, len(m))
		for k := range m {
			w.headers = append(w.headers, k)
		}
	}

	row := make([]string, len(w.headers))
	for i, header := range w.headers {
		if val, ok := m[header]; ok {
			row[i] = fmt.Sprint(val)
		}
	}
	return row, nil
}

// parseStruct 解析结构体类型的记录
func (w *BaseWriter) parseStruct(record any) ([]string, error) {
	v := reflect.ValueOf(record)
	if v.Kind() == reflect.Ptr {
		v = v.Elem()
	}
	if v.Kind() != reflect.Struct {
		return nil, fmt.Errorf("record must be a struct or map[string]any, got %T", record)
	}

	t := v.Type()
	if len(w.headers) == 0 {
		// 第一行数据，解析struct标签设置表头
		w.headers = make([]string, 0)
		for i := 0; i < t.NumField(); i++ {
			field := t.Field(i)
			tag := field.Tag.Get("table")
			if tag == "-" {
				continue
			}
			if tag == "" {
				tag = field.Name
			}
			w.headers = append(w.headers, tag)
		}
	}

	row := make([]string, len(w.headers))
	headerIndex := 0
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		tag := field.Tag.Get("table")
		if tag == "-" {
			continue
		}
		if headerIndex < len(row) {
			row[headerIndex] = fmt.Sprint(v.Field(i).Interface())
			headerIndex++
		}
	}
	return row, nil
}

// Close 完成写入
func (w *XlsxWriter) Close() error {
	// 写入表头
	for i, header := range w.headers {
		cell, err := excelize.CoordinatesToCellName(i+1, 1)
		if err != nil {
			return err
		}
		w.file.SetCellValue("Sheet1", cell, header)
	}

	// 写入数据
	for rowIndex, row := range w.data {
		for colIndex, value := range row {
			cell, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+2)
			if err != nil {
				return err
			}
			w.file.SetCellValue("Sheet1", cell, value)
		}
	}

	return w.file.SaveAs(w.filePath)
}

// Close 完成写入
func (w *CsvWriter) Close() error {
	writer := csv.NewWriter(w.file)
	defer w.file.Close()

	// 写入表头
	if err := writer.Write(w.headers); err != nil {
		return err
	}

	// 写入数据
	for _, row := range w.data {
		if err := writer.Write(row); err != nil {
			return err
		}
	}

	writer.Flush()
	return writer.Error()
}
