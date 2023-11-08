package excelplus

import (
	"github.com/ego-component/eos"
	"github.com/gotomicro/ego/core/elog"
	"github.com/xuri/excelize/v2"
)

type config struct {
	DefaultSheetName string
	EnableUpload     bool
}

func defaultConfig() *config {
	return &config{
		DefaultSheetName: "Sheet1",
		EnableUpload:     false,
	}
}

type Container struct {
	config *config
	s3     *eos.Component
}

// Load 第一个sheetname名称，默认Sheet1
func Load() *Container {
	return &Container{
		config: defaultConfig(),
	}
}

// Build 第一个sheetName名称，默认Sheet1
func (c *Container) Build(options ...Option) *Component {
	for _, option := range options {
		option(c)
	}

	// 如果开启了上传，那么就要判断s3是否为空
	if c.config.EnableUpload && c.s3 == nil {
		elog.Panic("s3 component is nil")
		return nil
	}

	exFile := excelize.NewFile()
	err := exFile.SetSheetName(exFile.GetSheetName(0), c.config.DefaultSheetName)
	if err != nil {
		elog.Panic("set sheet name error", elog.FieldErr(err))
		return nil
	}
	return &Component{
		File:   exFile,
		config: c.config,
		s3:     c.s3,
	}
}
