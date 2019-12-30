// #cgo CFLAGS: -O3
// #cgo CXXFLAGS: -O3
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
	mainwin = ui.NewWindow("Barbara's Tool", 700, 100, true)

	mainwin.OnClosing(func(*ui.Window) bool {
		ui.Quit()
		return true
	})

	ui.OnShouldQuit(func() bool {
		mainwin.Destroy()
		return true
	})

	// Controls
	vendorEntry := ui.NewEntry()
	vendorEntry.SetReadOnly(true)
	bidEntry := ui.NewEntry()
	bidEntry.SetReadOnly(true)
	buttonOpenFile := ui.NewButton("Open Files")
	buttonProcess := ui.NewButton("Process")

	// Filename form
	form := ui.NewForm()
	form.SetPadded(true)
	form.Append("", ui.NewLabel("Search Price (H) & Lead time (L) base on P/N (C)"), false)
	form.Append("Vendor", vendorEntry, false)
	form.Append("", ui.NewLabel("Fill in Price (K) & Lead time (P) base on P/N (F)"), false)
	form.Append("Bid", bidEntry, false)

	// Buttons
	buttonBox := ui.NewHorizontalBox()
	buttonBox.SetPadded(true)
	buttonBox.Append(ui.NewLabel(""), true)
	buttonBox.Append(ui.NewLabel("Cheers Barbara!"), false)
	buttonBox.Append(buttonOpenFile, false)
	buttonBox.Append(buttonProcess, false)

	// Main layout
	box := ui.NewVerticalBox()
	box.SetPadded(true)
	box.Append(form, false)
	box.Append(buttonBox, false)

	// Main window
	mainwin.SetChild(box)
	mainwin.SetMargined(true)
	mainwin.Show()

	// Open file button click event
	buttonOpenFile.OnClicked(func(*ui.Button) {
		vendorEntry.SetText(ui.OpenFile(mainwin))
		bidEntry.SetText(ui.OpenFile(mainwin))
	})

	// Process button click event
	buttonProcess.OnClicked(func(*ui.Button) {
		// Check vendor filename
		vendorFilename := strings.TrimSpace(vendorEntry.Text())
		if len(vendorFilename) == 0 {
			ui.MsgBox(mainwin, "Error", "Vendor filename cannot be empty.")
			return
		}
		if !strings.HasSuffix(vendorFilename, ".xlsx") {
			ui.MsgBox(mainwin, "Error", "Only support xlsx format :)")
			return
		}

		// Check bid filename
		bidFilename := strings.TrimSpace(bidEntry.Text())
		if len(bidFilename) == 0 {
			ui.MsgBox(mainwin, "Error", "Bid filename cannot be empty.")
			return
		}
		if !strings.HasSuffix(bidFilename, ".xlsx") {
			ui.MsgBox(mainwin, "Error", "Only support xlsx format :)")
			return
		}

		// Open vendor file
		vendorFile, err := excelize.OpenFile(vendorFilename)
		if err != nil {
			ui.MsgBox(mainwin, "Error", err.Error())
			return
		}

		// Open bid file
		bidFile, err := excelize.OpenFile(bidFilename)
		if err != nil {
			ui.MsgBox(mainwin, "Error", err.Error())
			return
		}

		// Make sure vendor file has at least 1 data sheet
		sheets := vendorFile.GetSheetMap()
		if len(sheets) == 0 {
			ui.MsgBox(mainwin, "Error", "The vendor file is empty!")
			return
		}

		// Create data maps by reading all vendor file sheet(s)
		// - PN        - Col C - "型号（P/N)"
		// - Price     - Col H - "合计（不含税）"
		// - Lead time - Col L - "货期"
		priceMap := make(map[string]string)
		leadTimeMap := make(map[string]string)
		for _, sheetName := range sheets {
			rows, err := vendorFile.GetRows(sheetName)
			if err != nil {
				ui.MsgBox(mainwin, "Error", err.Error())
				continue
			}

			for i := 0; i < len(rows); i++ {
				// Make sure PN & price exists
				if len(rows[i]) < 'H'-'A'+1 {
					continue
				}

				// Check PN
				pn := strings.ToLower(strings.TrimSpace(rows[i]['C'-'A']))
				if len(pn) == 0 {
					continue
				}

				// Check and set price
				price := strings.TrimSpace(rows[i]['H'-'A'])
				if _, err := strconv.ParseFloat(price, 64); err == nil {
					priceMap[pn] = price
				}

				// Make sure lead time exists
				if len(rows[i]) < 12 {
					continue
				}

				// Check and set lead time
				leadTime := strings.TrimSpace(rows[i]['L'-'A'])
				if len(leadTime) != 0 {
					leadTime = strings.ReplaceAll(leadTime, "周", " wks")
					leadTime = strings.ReplaceAll(leadTime, "现货", "In stock")
					leadTimeMap[pn] = leadTime
				}
			}
		}

		// Stop if both price map and lead time map are empty
		if len(priceMap) == 0 && len(leadTimeMap) == 0 {
			ui.MsgBox(mainwin, "Error", "No price or lead time data found from vendor file!")
			return
		}

		// Make sure bid file has at least 1 data sheet
		sheets = bidFile.GetSheetMap()
		if len(sheets) == 0 {
			ui.MsgBox(mainwin, "Error", "The bid excel is empty!")
			return
		}

		// Backup bid file
		bidFile.SaveAs("./backup." + time.Now().Format("20060102.150405.999999") + ".xlsx")

		// Update bid file
		pnFoundMap := make(map[string]bool)
		priceUpdateCounter := 0
		leadTimeUpdatecounter := 0
		for _, sheetName := range sheets {
			rows, err := bidFile.GetRows(sheetName)
			if err != nil {
				ui.MsgBox(mainwin, "Error", err.Error())
				continue
			}

			// Update bid file
			// - PN        - Col F - International Part No.
			// - Price     - Col K - Unit Price CNY
			// - Lead time - Col P - leadtime
			for i := 0; i < len(rows); i++ {
				// Make sure F exists
				length := len(rows[i])
				if length < 'F'-'A'+1 {
					continue
				}

				// Check PN
				pn := strings.ToLower(strings.TrimSpace(rows[i]['F'-'A']))
				if len(pn) == 0 {
					continue
				}

				// Update price
				if price, ok := priceMap[pn]; ok {
					pnFoundMap[pn] = true
					if length < 11 || price != rows[i]['K'-'A'] {
						if cellName, err := excelize.CoordinatesToCellName('K'-'A'+1, i+1); err == nil {
							bidFile.SetCellDefault(sheetName, cellName, price)
							priceUpdateCounter++
						}
					}
				}

				// Update lead time
				if leadTime, ok := leadTimeMap[pn]; ok {
					pnFoundMap[pn] = true
					if length < 16 || leadTime != rows[i]['P'-'A'] {
						if cellName, err := excelize.CoordinatesToCellName('P'-'A'+1, i+1); err == nil {
							bidFile.SetCellStr(sheetName, cellName, leadTime)
							leadTimeUpdatecounter++
						}
					}
				}
			}
		}

		// Save bid file if updated
		if priceUpdateCounter != 0 || leadTimeUpdatecounter != 0 {
			if err := bidFile.Save(); err != nil {
				ui.MsgBox(mainwin, "Error", "Error save bid file.\n"+err.Error())
				return
			}
		}

		// Done
		ui.MsgBox(mainwin, "Done!",
			fmt.Sprintf("%d matching PN(s) found from vendor file.\n"+
				"%d price cell(s) updated.\n"+
				"%d lead time cell(s) updated.",
				len(pnFoundMap), priceUpdateCounter, leadTimeUpdatecounter))
	})
}

func main() {
	ui.Main(setupUI)
}
