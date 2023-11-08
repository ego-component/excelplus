## 背景
excelize 每次需要传入[]any数据，这样会导致写入excel非常麻烦

## 解决方案
需要采用go的struct tag，将struct转换成[]any数据
目前是实验性质

```go
type SummaryBody struct {
    SheetName        string `excel:"Sheet名称"  excel_width:"20"`
    AvgIops          int64  `excel:"Avg iops"`
    AvgSpeed         string `excel:"Avg speed"`
    TotalWritesCount int    `excel:"Total writes count"`
    TotalWritesSize  uint64 `excel:"Total writes (MB)"`
}

//  创建一个 goexcel 实例
excelFile := goexcel.Load().Build(moexcel.WithDefaultSheetName(firstSheet))

// 创建一个 Sheet Header
moSummarySheet := excelFile.MoNewSheet(firstSheet, SummaryBody{})

// 设置 Sheet 的内容
moSummarySheet.SetRow(SummaryBody{
	xxx,
})

excelFile.SaveAs(fileName)
```
