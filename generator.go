package main

import (
	"crypto/rand"
	"errors"
	"log"
	"math/big"
	"os"
	"strconv"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

func generatePassword(length int, useLower, useUpper, useDigits bool) (string, error) {
	var charset string
	if useLower {
		charset += "abcdefghijklmnopqrstuvwxyz"
	}
	if useUpper {
		charset += "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	}
	if useDigits {
		charset += "0123456789"
	}

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
	a := app.New()
	w := a.NewWindow("Password Generator")
	w.Resize(fyne.NewSize(500, 300))

	labelEntry := widget.NewEntry()
	labelEntry.SetPlaceHolder("Enter password label")
	lengthEntry := widget.NewEntry()
	lengthEntry.SetPlaceHolder("Enter the length")
	passwordLabel := widget.NewLabel("Generated Password")

	lowerCheck := widget.NewCheck("Include Lowercase", nil)
	lowerCheck.SetChecked(true)
	upperCheck := widget.NewCheck("Include Uppercase", nil)
	upperCheck.SetChecked(true)
	digitsCheck := widget.NewCheck("Include Digits", nil)
	digitsCheck.SetChecked(true)

	generateBtn := widget.NewButton("Generate", func() {
		length, err := strconv.Atoi(lengthEntry.Text)
		if err != nil || length <= 0 {
			dialog.ShowError(errors.New("Invalid length"), w)
			return
		}
		password, err := generatePassword(length, lowerCheck.Checked, upperCheck.Checked, digitsCheck.Checked)
		if err != nil {
			dialog.ShowError(err, w)
			return
		}
		passwordLabel.SetText(password)
	})

	saveBtn := widget.NewButton("Save to Excel", func() {
		if passwordLabel.Text == "" || labelEntry.Text == "" {
			dialog.ShowError(errors.New("Password or label cannot be empty"), w)
			return
		}
		saveToExcel(labelEntry.Text, passwordLabel.Text)
		dialog.ShowInformation("Saved to Excel", "Saved to Excel successfully", w)
	})

	content := container.NewVBox(
		labelEntry,
		lengthEntry,
		lowerCheck,
		upperCheck,
		digitsCheck,
		generateBtn,
		passwordLabel,
		saveBtn,
	)
	w.SetContent(content)
	w.ShowAndRun()
}

func saveToExcel(label, password string) {
	filePath := "generator.xlsx"
	f, err := excelize.OpenFile(filePath)
	if err != nil && !os.IsNotExist(err) {
		log.Fatal("Failed to open file:", err)
		return
	}
	if f == nil {
		f = excelize.NewFile()
		f.NewSheet("Sheet1")
		f.SetCellValue("Sheet1", "A1", "Label")
		f.SetCellValue("Sheet1", "B1", "Password")
	}

	rows, _ := f.GetRows("Sheet1")
	row := len(rows) + 1
	cellA := "A" + strconv.Itoa(row)
	cellB := "B" + strconv.Itoa(row)
	f.SetCellValue("Sheet1", cellA, label)
	f.SetCellValue("Sheet1", cellB, password)

	if err := f.SaveAs(filePath); err != nil {
		log.Fatal("Failed to save file:", err)
	}
}
