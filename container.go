package goexcel

import (
	"github.com/ego-component/eos"
	"github.com/gotomicro/ego/core/elog"
	"github.com/xuri/excelize/v2"
)

type config struct {
	DefaultSheetName string
}

func defaultConfig() *config {
	return &config{
		DefaultSheetName: "Sheet1",
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
	exFile := excelize.NewFile()
	err := exFile.SetSheetName(exFile.GetSheetName(0), c.config.DefaultSheetName)
	if err != nil {
		elog.Panic("set sheet name error", elog.FieldErr(err))
		return nil
	}
	return &Component{
		File: exFile,
	}
}
