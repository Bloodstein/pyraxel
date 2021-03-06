package factory

import (
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/Bloodstein/pyraxel/models"
	"github.com/google/uuid"
)

type ExcelFactory struct{}

func (this *ExcelFactory) guid() string {
	uuidWithHyphen := uuid.New()
	fmt.Println(uuidWithHyphen)
	return uuidWithHyphen.String()
}

func (this *ExcelFactory) getColumns() []string {
	alpha := "A_B_C_D_E_F_G_H_I_J_K_L_M_N_O_P_Q_R_S_T_U_V_W_X_Y_Z"
	letters := strings.Split(alpha, "_")
	for _, a := range letters {
		for _, b := range letters {
			alpha = strings.Join([]string{alpha, strings.Join([]string{a, b}, "")}, "_")
		}
	}

	return strings.Split(alpha, "_")
}

func (this *ExcelFactory) generate(request models.ExcelRequest) string {

	f := excelize.NewFile()

	if _, err := os.Stat("/result"); os.IsNotExist(err) {
		os.Mkdir("./result", os.ModePerm)
	}

	var fileName string = fmt.Sprintf("%s.xlsx", this.guid())

	log.Printf("Name of file: %s\r\n", fileName)

	if len(request.Data.Simple) > 0 {
		log.Println("Get a simple data. Start fill it.")
		for _, cell := range request.Data.Simple {
			f.SetCellValue("Sheet1", cell.Address, cell.Value)
		}
	}

	columns := this.getColumns()

	log.Printf("Columns is got. Count: %s\r\n", fmt.Sprint(len(columns)))

	boldStyleId, _ := f.NewStyle(`{"font":{"bold":true}}`)

	log.Println("Start create a report's header")

	f.SetCellValue("Sheet1", strings.Join([]string{"A", fmt.Sprint(request.Params.Header.StartRow)}, ""), "#")

	for key, title := range request.Params.Header.Columns {
		cell := strings.Join([]string{columns[key+1], fmt.Sprint(request.Params.Header.StartRow)}, "")
		f.SetCellValue("Sheet1", cell, title)

		if request.Params.Header.Bold == true {
			f.SetCellStyle("Sheet1", cell, cell, boldStyleId)
		}
	}

	maxColumnNumber := len(request.Params.Header.Columns)

	if request.Params.Header.Filter == true {
		log.Println("The filter is need for report. Let's create it.")
		startCell := strings.Join([]string{"A", fmt.Sprint(request.Params.Header.StartRow)}, "")
		endCell := strings.Join([]string{columns[maxColumnNumber], fmt.Sprint(request.Params.Header.StartRow)}, "")
		f.AutoFilter("Sheet1", startCell, endCell, "")
	}

	dataStartRow := request.Params.Header.StartRow + 1
	log.Printf("Start row: %s\r\n", fmt.Sprint(dataStartRow))

	log.Printf("Start to fill report to table data. Count: %s", fmt.Sprint(len(request.Data.Table)))

	for index, row := range request.Data.Table {
		f.SetCellValue("Sheet1", strings.Join([]string{"A", fmt.Sprint(dataStartRow)}, ""), index+1)
		for key, value := range row {
			f.SetCellValue("Sheet1", strings.Join([]string{columns[key+1], fmt.Sprint(dataStartRow)}, ""), value)
		}
		dataStartRow++
	}

	log.Println("Filling a report table data was end")

	if err := f.SaveAs(fmt.Sprintf("./result/%s", fileName)); err != nil {
		log.Fatalf("Error was occured while saving an Excel file: %s", err.Error())
	}

	return fileName
}

func NewFactory() func(models.ExcelRequest) string {
	f := ExcelFactory{}
	return f.generate
}
