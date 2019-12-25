package main

import (
	"fmt"
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/andlabs/ui"
	_ "github.com/andlabs/ui/winmanifest"
)

var mainwin *ui.Window

func setupUI() {
	mainwin = ui.NewWindow("Barbara's Tool", 640, 100, true)
	mainwin.OnClosing(func(*ui.Window) bool {
		ui.Quit()
		return true
	})
	ui.OnShouldQuit(func() bool {
		mainwin.Destroy()
		return true
	})

	// Main layout
	box := ui.NewVerticalBox()
	box.SetPadded(true)

	// Filename form
	form := ui.NewForm()
	form.SetPadded(true)
	input := ui.NewEntry()
	input.SetReadOnly(true)
	output := ui.NewEntry()
	output.SetReadOnly(true)
	form.Append("", ui.NewLabel("Search Price (H) & Lead time (L) base on P/N (C)"), false)
	form.Append("Vendor", input, false)
	form.Append("", ui.NewLabel("Fill in Price (K) & Lead time (P) base on P/N (F)"), false)
	form.Append("Bid", output, false)
	box.Append(form, false)

	// Buttons
	buttonBox := ui.NewHorizontalBox()
	buttonBox.SetPadded(true)
	buttonOpenFile := ui.NewButton("Open Files")
	buttonProcess := ui.NewButton("Process")
	buttonBox.Append(ui.NewLabel(""), true)
	buttonBox.Append(ui.NewLabel("Hey Barbara! Calm down..."), false)
	buttonBox.Append(buttonOpenFile, false)
	buttonBox.Append(buttonProcess, false)
	box.Append(buttonBox, false)

	// Main window
	mainwin.SetChild(box)
	mainwin.SetMargined(true)
	mainwin.Show()

	// Open file button click event
	buttonOpenFile.OnClicked(func(*ui.Button) {
		input.SetText(ui.OpenFile(mainwin))
		output.SetText(ui.OpenFile(mainwin))
	})

	// Process button click event
	buttonProcess.OnClicked(func(*ui.Button) {
		// Check input filename
		inputFilename := strings.TrimSpace(input.Text())
		if len(inputFilename) == 0 {
			ui.MsgBox(mainwin, "Error", "Input filename cannot be empty.")
			return
		}
		if !strings.HasSuffix(inputFilename, ".xlsx") {
			ui.MsgBox(mainwin, "Error", "Only support xlsx format :)")
			return
		}

		// Check output filename
		outputFilename := strings.TrimSpace(output.Text())
		if len(outputFilename) == 0 {
			ui.MsgBox(mainwin, "Error", "Output filename cannot be empty.")
			return
		}
		if !strings.HasSuffix(outputFilename, ".xlsx") {
			ui.MsgBox(mainwin, "Error", "Only support xlsx format :)")
			return
		}

		// Open input file
		inputFile, err := excelize.OpenFile(inputFilename)
		if err != nil {
			ui.MsgBox(mainwin, "Error", err.Error())
			return
		}

		// Open output file
		outputFile, err := excelize.OpenFile(outputFilename)
		if err != nil {
			ui.MsgBox(mainwin, "Error", err.Error())
			return
		}

		// Make sure input file has at least 1 data sheet
		sheets := inputFile.GetSheetMap()
		if len(sheets) == 0 {
			ui.MsgBox(mainwin, "Error", "The price excel has no sheet!\n"+err.Error())
			return
		}

		// Create data maps by reading all input file sheet(s)
		// - PN        - Col C - "型号（P/N)"
		// - Price     - Col H - "合计（不含税）"
		// - Lead time - Col L - "货期"
		priceMap := make(map[string]string)
		leadTimeMap := make(map[string]string)
		for _, sheetName := range sheets {
			rows, err := inputFile.GetRows(sheetName)
			if err != nil {
				ui.MsgBox(mainwin, "Error", err.Error())
				continue
			}

			for i := 0; i < len(rows); i++ {
				// Make sure P/N & Price exists
				if len(rows[i]) < 8 {
					continue
				}

				// Check PN
				pn := strings.ToLower(strings.TrimSpace(rows[i][2]))
				if len(pn) == 0 {
					continue
				}

				// Check and set price
				price := strings.TrimSpace(rows[i][7])
				if _, err := strconv.ParseFloat(price, 64); err == nil {
					priceMap[pn] = price
				}

				// Make lead time exists
				if len(rows[i]) < 12 {
					continue
				}

				// Check and set leadtime
				leadTime := strings.TrimSpace(rows[i][11])
				if len(leadTime) != 0 {
					leadTime = strings.ReplaceAll(leadTime, "周", " wks")
					leadTime = strings.ReplaceAll(leadTime, "现货", "In stock")
					leadTimeMap[pn] = leadTime
				}
			}
		}

		// Stop if both price map and lead time map are empty
		if len(priceMap) == 0 && len(leadTimeMap) == 0 {
			ui.MsgBox(mainwin, "Error", "No price or lead time data found!")
			return
		}

		// Make sure input file has at least 1 data sheet
		sheets = outputFile.GetSheetMap()
		if len(sheets) == 0 {
			ui.MsgBox(mainwin, "Error", "The output excel has no sheet!\n"+err.Error())
			return
		}

		// Backup output file
		outputFile.SaveAs("./backup." + time.Now().Format("20060102-150405-0700") + ".xlsx")

		// Update output file
		updateCounter := 0
		for _, sheetName := range sheets {
			rows, err := outputFile.GetRows(sheetName)
			if err != nil {
				ui.MsgBox(mainwin, "Error", err.Error())
				continue
			}

			// Write output file
			// - PN        - Col F - International Part No.
			// - Price     - Col K - Unit Price CNY
			// - Lead time - Col P - leadtime
			for i := 0; i < len(rows); i++ {
				// Make sure F exists
				if len(rows[i]) < 6 {
					continue
				}

				// Check PN
				pn := strings.ToLower(strings.TrimSpace(rows[i][5]))
				if len(pn) == 0 {
					continue
				}

				// Update price
				if price, ok := priceMap[pn]; ok {
					if cellName, err := excelize.CoordinatesToCellName(11, i+1); err == nil {
						outputFile.SetCellDefault(sheetName, cellName, price)
						updateCounter++
					}
				}

				// Update lead time
				if leadtime, ok := leadTimeMap[pn]; ok {
					if cellName, err := excelize.CoordinatesToCellName(16, i+1); err == nil {
						outputFile.SetCellStr(sheetName, cellName, leadtime)
						// outputFile.SetCellDefault(sheetName, cellName, leadtime)
						updateCounter++
					}
				}
			}
		}

		// Nothing updated
		if updateCounter == 0 {
			ui.MsgBox(mainwin, "Error", "Nothing updated.")
			return
		}

		// Save file
		if err := outputFile.Save(); err != nil {
			ui.MsgBox(mainwin, "Error", "Save output excel.\n"+err.Error())
			return
		}

		// Done
		ui.MsgBox(mainwin, "Complete", fmt.Sprintf("Done! %d cells updated!", updateCounter))
	})
}

func main() {
	ui.Main(setupUI)
}
