package main

import (
	"encoding/csv"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

func main() {
	excelFilePath := "generator.xlsx"

	csvFilePath := "generator.csv"

	xlsx, err := excelize.OpenFile(excelFilePath)
	if err != nil {
		log.Fatal("Failed to open Excel file:", err)
	}

	csvFile, err := os.Create(csvFilePath)
	if err != nil {
		log.Fatal("Failed to create CSV file:", err)
	}
	defer csvFile.Close()

	sheetName := "Sheet1"

	rows, err := xlsx.GetRows(sheetName)
	if err != nil {
		log.Fatal("Failed to read Excel sheet:", err)
	}

	csvWriter := csv.NewWriter(csvFile)
	defer csvWriter.Flush()

	for _, row := range rows {
		if err := csvWriter.Write(row); err != nil {
			log.Fatal("Failed to write to CSV file:", err)
		}
	}

	log.Println("CSV file saved successfully:", csvFilePath)
}
