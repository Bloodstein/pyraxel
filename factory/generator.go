package factory

import (
	"fmt"
	"log"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/google/uuid"
)

type ExcelFactory struct {
}

func (this *ExcelFactory) guid() string {
	uuidWithHyphen := uuid.New()
	fmt.Println(uuidWithHyphen)
	return uuidWithHyphen.String()
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

func NewFactory() func([]string) string {
	f := ExcelFactory{}
	return f.generate
}
