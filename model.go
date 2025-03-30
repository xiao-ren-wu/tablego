package tablego

import (
	"github.com/spf13/cast"
)

// Item 单元格
type Item struct {
	// 行索引
	RowIndex int
	// 列索引
	CellIndex int
	// 键
	Key string
	// 值
	Value string
}

// Int64Value 单元格内容转换成int64
func (i *Item) Int64Value() (int64, error) {
	return cast.ToInt64E(i.Value)
}

// IntValue 单元格内容转换成int
func (i *Item) IntValue() (int, error) {
	return cast.ToIntE(i.Value)
}

// Float64Value 单元格内容转换成float64
func (i *Item) Float64Value() (float64, error) {
	return cast.ToFloat64E(i.Value)
}
