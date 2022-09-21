package excelbuddy

import (
	"github.com/xuri/excelize/v2"
	"io"
	"reflect"
)

type Options struct {
	SheetName string
	Columns   map[string]Column
}

type Validator interface {
	Validate(v string) error
}

type Column struct {
	Validators []Validator
}

type Assist struct {
	filename        string
	options         Options
	cellError       map[string]string
	columnConfigMap map[string]Column
	f               *excelize.File
	reader          io.Reader
	mode            string
}

type ColReflect struct {
	ColName     string
	ColIndex    int
	FieldIndex  int
	StructField reflect.StructField
}
