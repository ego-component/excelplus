package excelplus

import (
	"bytes"
	"context"
	"fmt"
	"os"
	"strings"

	"github.com/ego-component/eos"
	"github.com/gotomicro/ego/core/elog"
	"github.com/spf13/cast"
	"github.com/xuri/excelize/v2"
)

type Component struct {
	*excelize.File
	config *config
	s3     *eos.Component
	store  map[string]*Sheet
}

// ABC 只支持一部分列，如果超过26列，需要修改这里
const ABC = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

// NewSheet 新建sheet，同时所有操作都在这里面操作
func (c *Component) NewSheet(sheetName string, rowStruct any) (*Sheet, error) {
	oneSheet,flag := c.store[sheetName]
	if flag {
		return oneSheet,nil
	}
	
	reflectType, _, err := checkFields(rowStruct)
	if err != nil {
		panic(err)
	}
	// 通过标签名，获得excel的header
	headers := make([]string, 0)
	for i := 0; i < reflectType.NumField(); i++ {
		field := reflectType.Field(i)
		if field.Tag.Get("excel") != "" {
			headers = append(headers, field.Tag.Get("excel"))
		}
	}

	if len(headers) == 0 {
		return nil, fmt.Errorf("sheetHeader为空, 因为struct里面没有excel tag, sheetName是" + sheetName)
	}

	_, err = c.File.NewSheet(sheetName)
	if err != nil {
		return nil, fmt.Errorf("创建sheet失败, sheetName是"+sheetName+", err: %w", err)
	}

	for i := 0; i < reflectType.NumField(); i++ {
		field := reflectType.Field(i)
		if field.Tag.Get("excel_width") != "" {
			err = c.File.SetColWidth(sheetName, string(ABC[i]), string(ABC[i]), cast.ToFloat64(field.Tag.Get("excel_width")))
			if err != nil {
				return nil, fmt.Errorf("设置sheet宽度失败, sheetName是"+sheetName+", column是"+string(ABC[i])+", err: %w", err)
			}
		}

	}
	// 写入header
	err = c.File.SetSheetRow(sheetName, "A1", &headers)
	if err != nil {
		return nil, fmt.Errorf("写入sheetHeader失败, sheetName是"+sheetName+", err: %w", err)
	}

	err = c.File.SetCellStyle(sheetName, "A1", string(ABC[len(headers)-1])+"1", c.GetHeaderStyle())
	if err != nil {
		return nil, fmt.Errorf("设置sheetHeader默认样式失败, sheetName是"+sheetName+", err: %w", err)
	}
	sheetObj := &Sheet{
		SheetName: sheetName,
		Headers:   headers,
		cursor:    1,
		File:      c.File,
	}
	c.store[sheetName] = sheetObj
	return sheetObj, nil
}

// SaveAs 写入本地文件或者上传到s3
// 路径有 / 那么在本地和S3 自动创建文件夹
func (c *Component) SaveAs(ctx context.Context, fileName string) error {
	if strings.Contains(fileName, "/") {
		err := os.MkdirAll(fileName[:strings.LastIndex(fileName, "/")], 0755)
		if err != nil {
			return fmt.Errorf("创建文件夹失败, err: %w", err)
		}
	}
	err := c.File.SaveAs(fileName)
	if err != nil {
		return fmt.Errorf("保存到本地失败, err: %w", err)
	}

	if c.config.EnableUpload {
		fileBytes, err := os.ReadFile(fileName)
		if err != nil {
			elog.Panic("read excel error", elog.FieldErr(err), elog.FieldName(fileName))
		}
		// 上传到s3
		err = c.s3.Put(ctx, fileName, bytes.NewReader(fileBytes), nil)
		if err != nil {
			return fmt.Errorf("上传到s3失败, err: %w", err)
		}
	}
	return nil
}

func (c *Component) GetHeaderStyle() int {
	// 定义表头样式（通过结构体方式指定）
	headStyle, _ := c.File.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{
				Type:  "right",
				Color: "#000000",
				Style: 2,
			},
			{
				Type:  "left",
				Color: "#000000",
				Style: 2,
			},
			{
				Type:  "top",
				Color: "#A9A9A9",
				Style: 2,
			},
			{
				Type:  "bottom",
				Color: "#A9A9A9",
				Style: 2,
			},
		}, Fill: excelize.Fill{
			// gradient： 渐变色    pattern   填充图案
			// Pattern: 1,                   // 填充样式  当类型是 pattern 0-18 填充图案  1 实体填充
			// Color:   []string{"#FF0000"}, // 当Type = pattern 时，只有一个
			Type:  "gradient",
			Color: []string{"#A9A9A9", "#A9A9A9"},
			// 类型是 gradient 使用 0-5 横向(每种颜色横向分布) 纵向 对角向上 对角向下 有外向内 由内向外
			Shading: 1,
		}, Font: &excelize.Font{
			Bold: true,
			// Italic: false,
			// Underline: "single",
			Size:   14,
			Family: "宋体",
			// Strike:    true, // 删除线
			Color: "#FFFFFF",
		}, Alignment: &excelize.Alignment{
			// 水平对齐方式 center left right fill(填充) justify(两端对齐)  centerContinuous(跨列居中) distributed(分散对齐)
			Horizontal: "center",
			// 垂直对齐方式 center top  justify distributed
			Vertical: "center",
			// Indent:     1,        // 缩进  只要有值就变成了左对齐 + 缩进
			// TextRotation: 30, // 旋转
			// RelativeIndent:  10,   // 好像没啥用
			// ReadingOrder:    0,    // 不知道怎么设置
			// JustifyLastLine: true, // 两端分散对齐，只有 水平对齐 为 distributed 时 设置true 才有效
			// WrapText:        true, // 自动换行
			// ShrinkToFit:     true, // 缩小字体以填充单元格
		}, Protection: &excelize.Protection{
			Hidden: true,
			Locked: true,
		},
		// 内置的数字格式样式   0-638  常用的 0-58  配合lang使用，因为语言不同样式不同 具体的样式参照文档
		NumFmt: 0,
		NegRed: true,
	})
	return headStyle
}
