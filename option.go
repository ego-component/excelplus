package goexcel

import (
	"github.com/ego-component/eos"
)

type Option func(o *Container)

func WithDefaultSheetName(defaultSheetName string) Option {
	return func(o *Container) {
		o.config.DefaultSheetName = defaultSheetName
	}
}

func WithS3(s3 *eos.Component) Option {
	return func(o *Container) {
		o.s3 = s3
	}
}
