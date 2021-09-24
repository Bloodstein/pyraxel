package factory

import (
	"fmt"
	"log"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/Bloodstein/pyraxel/models"
	"github.com/google/uuid"
)

type ExcelFactory struct {
}

func (this *ExcelFactory) guid() string {
	uuidWithHyphen := uuid.New()
	fmt.Println(uuidWithHyphen)
	return uuidWithHyphen.String()
}

func (this *ExcelFactory) getColumns() []string {
	alpha := strings.Split("ABCDEFGHIJKLMNOPQRSTUVWXYZ", "")
	for _, a := range alpha {
		for _, b := range alpha {
			alpha = append(alpha, strings.Join([]string{a, b}, ""))
		}
	}

	return alpha
}

func (this *ExcelFactory) generate(data []string) string {

	f := excelize.NewFile()

	for key, value := range data {
		f.SetCellValue("Sheet1", "A"+fmt.Sprint(key+1), value)
	}

	var fileName string = fmt.Sprintf("%s.xlsx", this.guid())

	if err := f.SaveAs(fileName); err != nil {
		log.Fatalf("Error was occured while saving an Excel file: %s", err.Error())
	}

	return fileName
}

func (this *ExcelFactory) hardGeneration(request models.ExcelRequest) string {

	f := excelize.NewFile()

	var fileName string = fmt.Sprintf("%s.xlsx", this.guid())

	if len(request.Data.Simple) > 0 {
		for _, cell := range request.Data.Simple {
			f.SetCellValue("Sheet1", cell.Address, cell.Value)
		}
	}

	columns := this.getColumns()

	boldStyleId, _ := f.NewStyle(`{"font":{"bold":true}}`)
	var maxColumnNumber int

	for key, title := range request.Params.Header.Columns {
		maxColumnNumber = key
		cell := strings.Join([]string{columns[key], fmt.Sprint(request.Params.Header.StartRow)}, "")
		f.SetCellValue("Sheet1", cell, title)

		if request.Params.Header.Bold == true {
			f.SetCellStyle("Sheet1", cell, cell, boldStyleId)
		}
	}

	if request.Params.Header.Filter == true {
		f.AutoFilter("Sheet1", "A1", strings.Join([]string{columns[maxColumnNumber], fmt.Sprint(request.Params.Header.StartRow)}, ""), "")
	}

	dataStartRow := request.Params.Header.StartRow + 1

	for index, row := range request.Data.Table {
		f.SetCellValue("Sheet1", strings.Join([]string{"A", fmt.Sprint(dataStartRow)}, ""), index+1)
		for key, value := range row {
			f.SetCellValue("Sheet1", strings.Join([]string{columns[key+1], fmt.Sprint(dataStartRow)}, ""), value)
		}
		dataStartRow++
	}

	// {
	// 	"params": {
	// 		"header": {
	//			"rownum": 2
	// 			"columns": ["DebtID", "ЛС"],
	// 			"bold": true,
	// 			"filter": true
	// 		},
	// 	}
	// 	"data": {
	// 		"simple": [
	// 			{
	// 				"address": "A1",
	// 				"value": "Some report"
	// 			}
	// 		],
	// 		"table": [
	// 			["123", "2800222"]
	// 		]
	// 	}
	// }

	if err := f.SaveAs(fileName); err != nil {
		log.Fatalf("Error was occured while saving an Excel file: %s", err.Error())
	}

	return fileName
}

func NewFactory() func(models.ExcelRequest) string {
	f := ExcelFactory{}
	return f.hardGeneration
}
