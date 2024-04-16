package main

import (
	"log"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func main() {
	filePath := "generator.xlsx"
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatal("Unable to open file:", err)
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		log.Fatal("Unable to retrieve rows:", err)
	}

	for i, row := range rows {
		if i == 0 {
			f.SetCellValue("Sheet1", "F1", "Safety Score")
		} else if len(row) > 1 {
			password := row[1]
			score := evaluatePasswordSafety(password)
			scoreCell := "F" + strconv.Itoa(i+1)
			f.SetCellValue("Sheet1", scoreCell, score)
		}
	}

	if err := f.Save(); err != nil {
		log.Fatal("Unable to save file:", err)
	}
	log.Println("Safety scores have been updated.")
}

func evaluatePasswordSafety(password string) int {
	score := 0
	hasUpper := false
	hasLower := false
	hasDigit := false
	hasSpecial := false

	for _, c := range password {
		switch {
		case 'A' <= c && c <= 'Z':
			hasUpper = true
		case 'a' <= c && c <= 'z':
			hasLower = true
		case '0' <= c && c <= '9':
			hasDigit = true
		case (c >= '!' && c <= '/') || (c >= ':' && c <= '@') || (c >= '[' && c <= '`') || (c >= '{' && c <= '~'):
			hasSpecial = true
		}
	}

	score = len(password) * 4
	if hasUpper {
		score += 2
	}
	if hasLower {
		score += 2
	}
	if hasDigit {
		score += 2
	}
	if hasSpecial {
		score += 2
	}

	return score
}
