package tablego

import (
	"encoding/csv"
	"errors"
	"os"
	"path/filepath"
	"strings"

	"github.com/spf13/cast"
	"github.com/xuri/excelize/v2"
)

// BaseReader 基础读取器
type BaseReader struct {
	filePath string
	data     [][]string
	headers  []string
}

// ToHeaderKeyMapList 将表格转换成map列表，每一行是一个map，key是列名，value是单元格
func (b *BaseReader) ToHeaderKeyMapList() ([]map[string]*Item, error) {
	result := make([]map[string]*Item, 0, len(b.data))
	for rowIndex, row := range b.data {
		rowMap := make(map[string]*Item)
		for cellIndex, cell := range row {
			if cellIndex >= len(b.headers) {
				break
			}
			rowMap[b.headers[cellIndex]] = &Item{
				RowIndex:  rowIndex,
				CellIndex: cellIndex,
				Key:       b.headers[cellIndex],
				Value:     cell,
			}
		}
		result = append(result, rowMap)
	}
	return result, nil
}

// ToIndexKeyMapList 将表格转换成map列表，每一行是一个map，key是列索引，value是单元格
func (b *BaseReader) ToIndexKeyMapList() ([]map[string]*Item, error) {
	result := make([]map[string]*Item, 0, len(b.data))
	for rowIndex, row := range b.data {
		rowMap := make(map[string]*Item)
		for cellIndex, cell := range row {
			key := cast.ToString(cellIndex)
			rowMap[key] = &Item{
				RowIndex:  rowIndex,
				CellIndex: cellIndex,
				Key:       key,
				Value:     cell,
			}
		}
		result = append(result, rowMap)
	}
	return result, nil
}

// OneCell 获取表格中某一列的所有单元格
func (b *BaseReader) OneCell(key string) ([]*Item, error) {
	var columnIndex int = -1
	for i, header := range b.headers {
		if header == key {
			columnIndex = i
			break
		}
	}
	if columnIndex == -1 {
		return nil, errors.New("column not found")
	}

	result := make([]*Item, 0, len(b.data))
	for rowIndex, row := range b.data {
		if columnIndex >= len(row) {
			continue
		}
		result = append(result, &Item{
			RowIndex:  rowIndex,
			CellIndex: columnIndex,
			Key:       key,
			Value:     row[columnIndex],
		})
	}
	return result, nil
}

// XlsxReader xlsx读取器
type XlsxReader struct {
	BaseReader
}

// ToHeaderKeyMapList 实现TableReader接口
func (x *XlsxReader) ToHeaderKeyMapList() ([]map[string]*Item, error) {
	return x.BaseReader.ToHeaderKeyMapList()
}

// ToIndexKeyMapList 实现TableReader接口
func (x *XlsxReader) ToIndexKeyMapList() ([]map[string]*Item, error) {
	return x.BaseReader.ToIndexKeyMapList()
}

// OneCell 实现TableReader接口
func (x *XlsxReader) OneCell(key string) ([]*Item, error) {
	return x.BaseReader.OneCell(key)
}

// CsvReader csv读取器
type CsvReader struct {
	BaseReader
}

// ToHeaderKeyMapList 实现TableReader接口
func (c *CsvReader) ToHeaderKeyMapList() ([]map[string]*Item, error) {
	return c.BaseReader.ToHeaderKeyMapList()
}

// ToIndexKeyMapList 实现TableReader接口
func (c *CsvReader) ToIndexKeyMapList() ([]map[string]*Item, error) {
	return c.BaseReader.ToIndexKeyMapList()
}

// OneCell 实现TableReader接口
func (c *CsvReader) OneCell(key string) ([]*Item, error) {
	return c.BaseReader.OneCell(key)
}

// NewTableReader 创建表格读取器
func NewTableReader(filePath string) (TableReader, error) {
	ext := strings.ToLower(filepath.Ext(filePath))
	switch ext {
	case ".xlsx":
		return newXlsxReader(filePath)
	case ".csv":
		return newCsvReader(filePath)
	default:
		return nil, errors.New("unsupported file type")
	}
}

// newXlsxReader 创建xlsx读取器
func newXlsxReader(filePath string) (*XlsxReader, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil, errors.New("no sheet found")
	}

	rows, err := f.GetRows(sheets[0])
	if err != nil {
		return nil, err
	}
	if len(rows) == 0 {
		return nil, errors.New("empty sheet")
	}

	return &XlsxReader{
		BaseReader: BaseReader{
			filePath: filePath,
			data:     rows[1:],
			headers:  rows[0],
		},
	}, nil
}

// newCsvReader 创建csv读取器
func newCsvReader(filePath string) (*CsvReader, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	reader := csv.NewReader(f)
	rows, err := reader.ReadAll()
	if err != nil {
		return nil, err
	}
	if len(rows) == 0 {
		return nil, errors.New("empty file")
	}

	return &CsvReader{
		BaseReader: BaseReader{
			filePath: filePath,
			data:     rows[1:],
			headers:  rows[0],
		},
	}, nil
}
