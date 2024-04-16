package main

import (
	"log"
	"strconv"

	"github.com/xuri/excelize/v2"
	"golang.org/x/crypto/bcrypt"
)

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

	xlsx.SetCellValue(sheetName, "H1", "Hashed")
	startRow := 1
	for i, row := range rows {
		if len(row) == 0 {
			continue
		}
		if i == 0 {
			continue
		}
		password := row[1]
		if password != "" {
			hashedPassword, err := bcrypt.GenerateFromPassword([]byte(password), bcrypt.DefaultCost)
			if err != nil {
				log.Printf("Failed to hash password at row %d: %v", i+1, err)
				continue
			}
			xlsx.SetCellValue(sheetName, "H"+strconv.Itoa(startRow+i), string(hashedPassword))
		}
	}

	if err := xlsx.SaveAs(excelFilePath); err != nil {
		log.Fatal("Failed to save Excel file:", err)
	}

	log.Println("Hashed passwords have been saved to Excel file:", excelFilePath)
}
