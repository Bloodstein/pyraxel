package handlers

import (
	"encoding/base64"
	"encoding/json"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"net/http"

	"github.com/Bloodstein/pyraxel/factory"
	"github.com/Bloodstein/pyraxel/models"
)

type FileResponse struct {
	FileName string `json:"fileName"`
	Content  string `json:"buffer"`
}

type Response struct {
	Status  string       `json:"status"`
	Message string       `json:"message"`
	File    FileResponse `json:"file"`
}

func NewHandler() func() {
	return handleRequest
}

func handleRequest() {
	http.HandleFunc("/generate-simple-excel", executeGeneration)

	if err := http.ListenAndServe(":8080", nil); err != nil {
		log.Fatalf("Error was occured while serving HTTP: %s", err.Error())
	}

	fmt.Println("Server has been start on localhost:8080. Try it!")
}

func executeGeneration(writer http.ResponseWriter, request *http.Request) {
	log.Println("Starting generation...")

	decoder := json.NewDecoder(request.Body)

	var requestParams models.ExcelRequest

	if err := decoder.Decode(&requestParams); err != nil {
		log.Fatalf("Error was occured while unmarchal request body: %s", err.Error())
	}

	method := factory.NewFactory()
	fileName := method(requestParams)

	log.Println("Generation is end!")

	fileContent, err := ioutil.ReadFile(fileName)

	if err != nil {
		log.Fatalf("Error was occured while open file: %s", err.Error())
	}

	res := Response{
		Status:  "ok",
		Message: "Generation was end. An Excel file was create successfully!",
		File: FileResponse{
			FileName: fileName,
			Content:  base64.StdEncoding.EncodeToString(fileContent),
		},
	}

	responseBytes, err := json.Marshal(res)

	if err != nil {
		log.Fatalf("Error was occured while marshaling response structure: %s", err.Error())
	}

	response := string(responseBytes)

	writer.WriteHeader(200)
	io.WriteString(writer, response)
}
