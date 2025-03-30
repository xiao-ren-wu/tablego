package main

import (
	"fmt"
	"log"

	"github.com/spf13/cast"

	"github.com/xiao-ren-wu/tablego"
)

// User 用户信息
type User struct {
	Name   string `table:"name"`
	Age    int    `table:"age"`
	Gender string `table:"gender"`
}

func main() {
	// 读取xlsx文件示例
	xlsxReader, err := tablego.NewTableReader("./testdata/users.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// 使用map方式读取
	fmt.Println("=== 使用map方式读取xlsx文件 ===")
	rows, err := xlsxReader.ToHeaderKeyMapList()
	if err != nil {
		log.Fatal(err)
	}
	for _, row := range rows {
		name := row["name"].Value
		age := cast.ToInt(row["age"].Value)
		gender := row["gender"].Value
		fmt.Printf("姓名: %s, 年龄: %d, 性别: %s\n", name, age, gender)
	}

	// 读取单列数据
	fmt.Println("\n=== 读取xlsx文件中的name列 ===")
	names, err := xlsxReader.OneCell("name")
	if err != nil {
		log.Fatal(err)
	}
	for _, name := range names {
		fmt.Printf("姓名: %s\n", name.Value)
	}

	// 写入xlsx文件示例
	xlsxWriter, err := tablego.NewTableWriter("./testdata/users_new.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	defer xlsxWriter.Close()

	// 使用struct方式写入
	fmt.Println("\n=== 使用struct方式写入xlsx文件 ===")
	users := []User{
		{Name: "张三", Age: 18, Gender: "男"},
		{Name: "李四", Age: 20, Gender: "女"},
	}
	for _, user := range users {
		if err := xlsxWriter.WriteLine(user); err != nil {
			log.Fatal(err)
		}
	}

	// 读取csv文件示例
	csvReader, err := tablego.NewTableReader("./testdata/users.csv")
	if err != nil {
		log.Fatal(err)
	}

	// 使用map方式读取
	fmt.Println("\n=== 使用map方式读取csv文件 ===")
	rows, err = csvReader.ToHeaderKeyMapList()
	if err != nil {
		log.Fatal(err)
	}
	for _, row := range rows {
		name := row["name"].Value
		age := cast.ToInt(row["age"].Value)
		gender := row["gender"].Value
		fmt.Printf("姓名: %s, 年龄: %d, 性别: %s\n", name, age, gender)
	}

	// 写入csv文件示例
	csvWriter, err := tablego.NewTableWriter("./testdata/users_new.csv")
	if err != nil {
		log.Fatal(err)
	}
	defer csvWriter.Close()

	// 使用struct方式写入
	fmt.Println("\n=== 使用struct方式写入csv文件 ===")
	for _, user := range users {
		if err := csvWriter.WriteLine(user); err != nil {
			log.Fatal(err)
		}
	}
}
