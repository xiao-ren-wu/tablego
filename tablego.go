package tablego

type TableReader interface {
	// 将表格转换成map列表，每一行是一个map，key是列名，value是单元格
	ToHeaderKeyMapList() ([]map[string]*Item, error)
	// 将表格转换成map列表，每一行是一个map，key是列索引，value是单元格
	ToIndexKeyMapList() ([]map[string]*Item, error)
	// 获取表格中某一列的所有单元格
	OneCell(key string) ([]*Item, error)
}

type TableWriter interface {
	// 追加写入一行
	WriteLine(record any) error
	// 完成写入
	Close() error
}
