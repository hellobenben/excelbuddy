package main

import (
	"fmt"
	"github.com/hellobenben/excelbuddy"
	"github.com/hellobenben/excelbuddy/validator"
	"os"
)

func main() {
	f, err := os.OpenFile("./example/demo.xlsx", os.O_RDWR, 0755)
	if err != nil {
		panic(err)
	}
	defer func() {
		f.Close()
	}()
	//assist, err := gopkg_excel.Open("./example/demo.xlsx")
	assist, err := excelbuddy.OpenReader(f)
	if err == nil {
		defer func() {
			assist.Close()
		}()
	}
	assist.SetColumnValidators("age", []excelbuddy.Validator{validator.RegExpValidator{Pattern: "^[1-9][1-9]$"}})
	assist.SetColumnValidators("email", []excelbuddy.Validator{validator.Required{}})
	assist.SetColumnValidators("中文", []excelbuddy.Validator{validator.Required{}})
	type Row struct {
		Name  string `excelbuddy:"name"`
		Age   int    `excelbuddy:"age"`
		Email string `excelbuddy:"email"`
		Ch    string `excelbuddy:"中文"`
	}
	var rows []Row
	err = assist.Scan(&rows)
	if err != nil {
		fmt.Println(err)
	}

	if assist.Validate() {
		assist.MarkError()
		err := assist.SaveAs("./example/demo_err.xlsx")
		fmt.Println(err)
	}

	fmt.Printf("%+v", rows)
}
