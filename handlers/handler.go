package handlers

import (
	"encoding/base64"
	"encoding/json"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"net/http"
	"strings"

	"github.com/Bloodstein/pyraxel/factory"
)

type Response struct {
	Status   string `json:"status"`
	FileName string `json:"fileName"`
	Content  string `json:"buffer"`
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

	method := factory.NewFactory()
	fileName := method(strings.Split(request.URL.Query().Get("values"), ","))

	log.Println("Generation is end!")

	fileContent, err := ioutil.ReadFile(fileName)

	if err != nil {
		log.Fatalf("Error was occured while open file: %s", err.Error())
	}

	res := Response{
		Status:   "ok",
		FileName: fileName,
		Content:  base64.StdEncoding.EncodeToString(fileContent),
	}

	bRes, err := json.Marshal(res)

	if err != nil {
		log.Fatalf("Error was occured while marshaling response structure: %s", err.Error())
	}

	response := fmt.Sprintf("{\"status\": \"ok\", \"message\": \"Generation was end. An Excel file was create successfully!\", \"response\": %s}", string(bRes))

	writer.WriteHeader(200)
	io.WriteString(writer, response)
}
