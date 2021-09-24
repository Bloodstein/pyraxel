package models

// Заголовок: текст, жирный, фильтр

type HeaderParams struct {
	StartRow int      `json:"startRow"`
	Columns  []string `json:"columns"`
	Bold     bool     `json:"bold"`
	Filter   bool     `json:"filter"`
}

type ExcelParams struct {
	Header HeaderParams `json:"header"`
}

type SimpleCell struct {
	Address string `json:"address"`
	Value   string `json:"value"`
}

type TableRow []string

type ExcelData struct {
	Simple []SimpleCell `json:"simple"`
	Table  []TableRow   `json:"table"`
}

type ExcelRequest struct {
	Params ExcelParams `json:"params"`
	Data   ExcelData   `json:"data"`
}
