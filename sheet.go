package excelplus

import (
	"fmt"
	"reflect"

	"github.com/xuri/excelize/v2"
)

type Sheet struct {
	SheetName string
	Headers   []string
	cursor    int
	File      *excelize.File
}

func (s *Sheet) SetRow(rowStruct any) error {
	s.cursor++
	rows := getRows(rowStruct, s.Headers)
	err := s.File.SetSheetRow(s.SheetName, fmt.Sprintf("A%v", s.cursor), &rows)
	if err != nil {
		return fmt.Errorf("set sheet row error: %w", err)
	}
	return nil
}

func getRows(row any, headers []string) []any {
	reflectType, reflectValue, err := checkFields(row)
	if err != nil {
		panic(err)
	}
	output := make([]any, 0)
	for _, value := range headers {
		flag := false
		for i := 0; i < reflectType.NumField(); i++ {
			field := reflectType.Field(i)
			// 当有excel标签，那么就插入数据
			if field.Tag.Get("excel") == value {
				flag = true
				output = append(output, reflectValue.Field(i).Interface())
				break
			}
		}
		// 如果为false，说明没找到，那么就插入空数据
		if !flag {
			output = append(output, "")
		}
	}
	return output
}

func checkFields(subject interface{}) (reflect.Type, reflect.Value, error) {
	typeOfSubject := reflect.TypeOf(subject)
	valueOfSubject := reflect.ValueOf(subject)

	//switch里面判断subject的类型，如果是结构体指针类型则做一系列转换，获取结构体类型
	switch typeOfSubject.Kind() {
	case reflect.Struct:
		break
	case reflect.Ptr: //如果是指针类型，则需要通过Elem()函数得到它的实际数据类型
		for typeOfSubject.Kind() == reflect.Ptr {
			typeOfSubject = typeOfSubject.Elem()
		}
		if typeOfSubject.Kind() != reflect.Struct { //如果实际数据类型不是结构体类型，则返回错误
			return nil, valueOfSubject, fmt.Errorf("error: subject can not be " + typeOfSubject.Kind().String())
		}
	default: //如果不是结构体类型也不是指针类型，则返回错误
		return nil, valueOfSubject, fmt.Errorf("error: subject can not be " + typeOfSubject.Kind().String())
	}
	return typeOfSubject, valueOfSubject, nil
}
