package main

import (
	"crypto/rand"
	"log"
	"math/big"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func generatePassword(length int) (string, error) {
	const charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
	password := make([]byte, length)
	for i := range password {
		index, err := rand.Int(rand.Reader, big.NewInt(int64(len(charset))))
		if err != nil {
			return "", err
		}
		password[i] = charset[index.Int64()]
	}
	return string(password), nil
}

func main() {
	excelFilePath := "generator.xlsx"

	xlsx, err := excelize.OpenFile(excelFilePath)
	if err != nil {
		log.Fatal("Failed to open Excel file:", err)
	}

	sheetName := "Sheet1"

	rows, err := xlsx.GetRows(sheetName)
	if err != nil {
		log.Fatal("Failed to read Excel sheet:", err)
	}

	startRow := 1
	for i, row := range rows {
		if len(row) == 0 {
			continue
		}
		if i == 0 {
			continue
		}
		oldPassword := row[1]
		password, err := generatePassword(12)
		if err != nil {
			log.Printf("Failed to generate password at row %d: %v", i+1, err)
			continue
		}
		xlsx.SetCellValue(sheetName, "P"+strconv.Itoa(startRow+i), oldPassword)
		xlsx.SetCellValue(sheetName, "B"+strconv.Itoa(startRow+i), password)
	}

	if err := xlsx.SaveAs(excelFilePath); err != nil {
		log.Fatal("Failed to save Excel file:", err)
	}

	log.Println("Passwords have been regenerated and saved to Excel file:", excelFilePath)
}
