package excelbuddy

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"reflect"
	"strconv"
)

func Open(filename string) (*Assist, error) {
	assist := &Assist{
		filename: filename,
		mode:     "file",
	}
	initAssist(assist)
	var err error
	assist.f, err = excelize.OpenFile(filename)
	if err != nil {
		return nil, err
	}
	return assist, nil
}

func OpenReader(r io.Reader) (*Assist, error) {
	assist := &Assist{
		reader: r,
		mode:   "stream",
	}
	initAssist(assist)
	var err error
	assist.f, err = excelize.OpenReader(r)
	if err != nil {
		return nil, err
	}
	return assist, nil
}

func initAssist(assist *Assist) {
	assist.cellError = make(map[string]string)
	assist.columnConfigMap = make(map[string]Column)
	assist.options.SheetName = "Sheet1"
}

func (a *Assist) Close() error {
	if a.f != nil {
		return a.f.Close()
	}
	return nil
}

func (a *Assist) Options(options Options) *Assist {
	a.options = options
	return a
}

func (a *Assist) SetColumnValidators(name string, validators []Validator) *Assist {
	if _, ok := a.columnConfigMap[name]; !ok {
		a.columnConfigMap[name] = Column{}
	}
	c := a.columnConfigMap[name]
	c.Validators = validators
	a.columnConfigMap[name] = c
	return a
}

func (a *Assist) Scan(dst interface{}) error {
	dstT := reflect.TypeOf(dst).Elem()
	if dstT.Kind() != reflect.Slice {
		return errors.New("dst must be a slice")
	}
	colMap := map[string]ColReflect{}
	itemT := dstT.Elem()
	for i := 0; i < itemT.NumField(); i++ {
		fieldType := itemT.Field(i)
		colName := fieldType.Tag.Get("excelbuddy")
		colMap[colName] = ColReflect{
			ColName:     colName,
			FieldIndex:  i,
			StructField: fieldType,
		}
	}

	// Get all the rows in the Sheet1.
	rows, err := a.f.GetRows(a.options.SheetName)
	if err != nil {
		return err
	}
	if len(rows) == 0 {
		return nil
	}
	for i, name := range rows[0] {
		if col, ok := colMap[name]; ok {
			col.ColIndex = i
			colMap[name] = col
		}
	}
	rowsSlice := reflect.MakeSlice(reflect.SliceOf(itemT), 0, 10)
	for rowIndex := 1; rowIndex < len(rows); rowIndex++ {
		row := rows[rowIndex]
		itemV := reflect.New(itemT)
		hasErr := false
		for _, col := range colMap {
			field := itemV.Elem().Field(col.FieldIndex)
			var v string
			if len(row) > col.ColIndex {
				v = row[col.ColIndex]
			}

			if field.Kind() == reflect.Int {
				intV, err := strconv.Atoi(v)
				if err != nil {
					a.addError(rowIndex+1, col.ColIndex+1, err.Error())
					hasErr = true
				}
				// validate
				for _, validator := range a.columnConfigMap[col.ColName].Validators {
					err := validator.Validate(v)
					if err != nil {
						a.addError(rowIndex+1, col.ColIndex+1, err.Error())
						hasErr = true
						break
					}
				}
				field.SetInt(int64(intV))
			} else {
				// validate
				for _, validator := range a.columnConfigMap[col.ColName].Validators {
					err := validator.Validate(v)
					if err != nil {
						a.addError(rowIndex+1, col.ColIndex+1, err.Error())
						hasErr = true
						break
					}
				}
				field.SetString(v)
			}
		}
		if hasErr {
			continue
		}
		rowsSlice = reflect.Append(rowsSlice, itemV.Elem())
	}
	dstV := reflect.ValueOf(dst).Elem()
	dstV.Set(rowsSlice)
	return nil
}

func (a *Assist) Validate() bool {
	return len(a.cellError) == 0
}

func (a *Assist) MarkError() {
	style, err := a.f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#FF0000"}, Pattern: 1},
	})
	if err != nil {
		fmt.Println(err)
	}
	for cell, msg := range a.cellError {
		_ = a.f.SetCellStyle(a.options.SheetName, cell, cell, style)
		_ = a.f.AddComment(a.options.SheetName, cell, fmt.Sprintf("{\"author\":\"Validator: \",\"text\":\"%s.\"}", msg))
	}
}

func (a *Assist) Save() error {
	return a.f.Save()
}

func (a *Assist) SaveAs(filename string) error {
	return a.f.SaveAs(filename)
}

func (a *Assist) addError(row int, col int, msg string) {
	colName, _ := excelize.ColumnNumberToName(col)
	cell := fmt.Sprintf("%s%d", colName, row)
	a.cellError[cell] = msg
}
