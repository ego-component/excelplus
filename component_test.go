package excelplus

import (
	"fmt"
	"testing"
)

type TestExcel struct {
	Name  string `excel:"姓名"`
	Age   int    `excel:"年龄"`
	Hello string
}

func Test_getRows(t *testing.T) {
	test := TestExcel{
		Name:  "我是",
		Age:   1,
		Hello: "123",
	}
	rows := getRows(test, []string{"姓名", "年龄"})
	fmt.Printf("rows--------------->"+"%+v\n", rows)
}
