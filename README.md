# TableGo

TableGo 是一个简单易用的 Go 语言表格处理库，支持读写 xlsx 和 csv 格式的文件。它提供了统一的接口，让你可以用相同的方式处理不同格式的表格文件。

## 特性

- 支持读写 xlsx 和 csv 格式文件
- 提供统一的 TableReader 和 TableWriter 接口
- 支持 map 和 struct 两种数据读写方式
- 使用 struct tag 映射表格列名

## 安装

```bash
go get github.com/tablego
```

## 快速开始

### 读取表格文件

```go
// 创建表格读取器
reader, err := tablego.NewTableReader("users.xlsx")
if err != nil {
    log.Fatal(err)
}

// 使用 map 方式读取
rows, err := reader.ToHeaderKeyMapList()
if err != nil {
    log.Fatal(err)
}
for _, row := range rows {
    name := row["name"].Value
    age := cast.ToInt(row["age"].Value)
    fmt.Printf("姓名: %s, 年龄: %d\n", name, age)
}

// 读取单列数据
names, err := reader.OneCell("name")
if err != nil {
    log.Fatal(err)
}
for _, name := range names {
    fmt.Printf("姓名: %s\n", name.Value)
}
```

### 写入表格文件

```go
// 定义数据结构
type User struct {
    Name   string `table:"name"`
    Age    int    `table:"age"`
    Gender string `table:"gender"`
}

// 创建表格写入器
writer, err := tablego.NewTableWriter("users.xlsx")
if err != nil {
    log.Fatal(err)
}
defer writer.Close()

// 使用 struct 方式写入
user := User{Name: "张三", Age: 18, Gender: "男"}
if err := writer.WriteLine(user); err != nil {
    log.Fatal(err)
}
```

## API 文档

### TableReader 接口

```go
type TableReader interface {
    // 将表格转换成 map 列表，每一行是一个 map，key 是列名，value 是单元格
    ToHeaderKeyMapList() ([]map[string]*Item, error)
    // 将表格转换成 map 列表，每一行是一个 map，key 是列索引，value 是单元格
    ToIndexKeyMapList() ([]map[string]*Item, error)
    // 获取表格中某一列的所有单元格
    OneCell(key string) ([]*Item, error)
}
```

### TableWriter 接口

```go
type TableWriter interface {
    // 追加写入一行
    WriteLine(record any) error
    // 完成写入
    Close() error
}
```

### struct tag 使用说明

在使用 struct 方式写入数据时，需要通过 `table` tag 指定字段对应的列名：

```go
type User struct {
    Name   string `table:"name"`   // 对应表格中的 name 列
    Age    int    `table:"age"`    // 对应表格中的 age 列
    Gender string `table:"gender"` // 对应表格中的 gender 列
}
```

## 完整示例

查看 [example](./example) 目录获取完整的使用示例。

## License

MIT License