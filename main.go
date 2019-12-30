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

// Open a xlsx file
func openExcelFile(must bool) (filename string) {
	for true {
		filename = ui.OpenFile(mainwin)

		// The open file dialog is cancelled
		if len(filename) == 0 {
			if must {
				ui.MsgBoxError(mainwin, "Error", "A file must be selected. Please try again.")
				continue
			} else {
				break
			}
		}

		// Not an xlsx file
		if !strings.HasSuffix(filename, ".xlsx") {
			ui.MsgBoxError(mainwin, "Error", "Only support xlsx format. Please try again.")
			continue
		}

		break
	}

	return
}

func setupUI() {
	mainwin = ui.NewWindow("Barbara's Tool", 700, 100, false)

	mainwin.OnClosing(func(*ui.Window) bool {
		ui.Quit()
		return true
	})

	ui.OnShouldQuit(func() bool {
		mainwin.Destroy()
		return true
	})

	// Controls
	entryBidFile := ui.NewEntry()
	entryBidFile.SetReadOnly(true)
	buttonAdd := ui.NewButton("Add Vendor File")
	buttonReset := ui.NewButton(" R ")
	buttonProcess := ui.NewButton("Process")

	// Add vendor file button
	boxAdd := ui.NewHorizontalBox()
	boxAdd.SetPadded(true)
	boxAdd.Append(ui.NewLabel(""), true)
	boxAdd.Append(buttonAdd, false)

	// Vendor files
	boxVendorFiles := ui.NewVerticalBox()
	boxVendorFiles.SetPadded(true)

	// Bid file box
	boxBidFile := ui.NewHorizontalBox()
	boxBidFile.SetPadded(true)
	boxBidFile.Append(entryBidFile, true)
	boxBidFile.Append(buttonReset, false)

	// Process button layout
	boxProcess := ui.NewHorizontalBox()
	boxProcess.SetPadded(true)
	boxProcess.Append(ui.NewLabel("Cheers Barbara!"), false)
	boxProcess.Append(ui.NewLabel(""), true)
	boxProcess.Append(buttonProcess, false)

	// Main layout
	box := ui.NewVerticalBox()
	box.SetPadded(true)
	box.Append(ui.NewLabel("Vendor file - Search Price (H) & Lead time (L) base on P/N (C)"), false)
	box.Append(boxVendorFiles, false)
	box.Append(ui.NewVerticalBox(), true)
	box.Append(boxAdd, false)
	box.Append(ui.NewHorizontalSeparator(), false)
	box.Append(ui.NewLabel("Bid file - Fill in Price (K) & Lead time (P) base on P/N (F)"), false)
	box.Append(boxBidFile, false)
	box.Append(ui.NewHorizontalSeparator(), false)
	box.Append(boxProcess, false)

	// Keep track of vendor files entries for removal
	vendorEntries := make([]*ui.Entry, 0)

	// Add a new vendor file
	addVendorFile := func() (added bool) {
		// Open file, the first file is required.
		filename := openExcelFile(len(vendorEntries) == 0)

		if len(filename) != 0 {
			// Add a new entry to UI
			entry := ui.NewEntry()
			entry.SetReadOnly(true)
			entry.SetText(filename)
			button := ui.NewButton(" - ")
			entryBox := ui.NewHorizontalBox()
			entryBox.SetPadded(true)
			entryBox.Append(entry, true)
			entryBox.Append(button, false)
			boxVendorFiles.Append(entryBox, false)

			// Append the entry
			vendorEntries = append(vendorEntries, entry)

			button.OnClicked(func(*ui.Button) {
				for i, e := range vendorEntries {
					if e == entry {
						// Destroy the file entry UI
						boxVendorFiles.Delete(i)
						entryBox.Destroy() // It cascades to all children

						// Remove the entry from the tracking slice
						vendorEntries = append(vendorEntries[:i], vendorEntries[i+1:]...)
						break
					}
				}
			})

			return true
		}

		return false
	}

	// Main window
	mainwin.SetChild(box)
	mainwin.SetMargined(true)
	mainwin.Show()

	// Open files
	ui.MsgBox(mainwin, "Message", "Open the vendor files.")
	for addVendorFile() { // Until cancelled
	}
	ui.MsgBox(mainwin, "Message", "Open the bid file.")
	entryBidFile.SetText(openExcelFile(true))

	// Add vendor file button click event
	buttonAdd.OnClicked(func(*ui.Button) {
		addVendorFile()
	})

	// Reset bid file
	buttonReset.OnClicked(func(*ui.Button) {
		filename := openExcelFile(false)
		if len(filename) != 0 {
			entryBidFile.SetText(filename)
		}
	})

	// Process button click event
	buttonProcess.OnClicked(func(*ui.Button) {
		// Create data maps by reading all vendor file sheet(s)
		// - PN        - Col C - "型号（P/N)"
		// - Price     - Col H - "合计（不含税）"
		// - Lead time - Col L - "货期"
		priceMap := make(map[string]string)
		leadTimeMap := make(map[string]string)
		for _, entry := range vendorEntries {
			// Open vendor file
			vendorFile, err := excelize.OpenFile(entry.Text())
			if err != nil {
				ui.MsgBoxError(mainwin, "Error", err.Error())
				return
			}

			for _, sheetName := range vendorFile.GetSheetMap() {
				rows, err := vendorFile.GetRows(sheetName)
				if err != nil {
					ui.MsgBoxError(mainwin, "Error", err.Error())
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
		}

		// Stop if both price map and lead time map are empty
		if len(priceMap) == 0 && len(leadTimeMap) == 0 {
			ui.MsgBoxError(mainwin, "Error", "No price or lead time data found from vendor file(s)!")
			return
		}

		// Open bid file
		bidFile, err := excelize.OpenFile(entryBidFile.Text())
		if err != nil {
			ui.MsgBoxError(mainwin, "Error", err.Error())
			return
		}

		// Backup bid file
		bidFile.SaveAs("./backup." + time.Now().Format("20060102.150405.999999") + ".xlsx")

		// Update bid file
		pnFoundMap := make(map[string]bool)
		priceUpdateCounter := 0
		leadTimeUpdatecounter := 0
		for _, sheetName := range bidFile.GetSheetMap() {
			rows, err := bidFile.GetRows(sheetName)
			if err != nil {
				ui.MsgBoxError(mainwin, "Error", err.Error())
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
				ui.MsgBoxError(mainwin, "Error", "Error save bid file.\n"+err.Error())
				return
			}
		}

		// Done
		ui.MsgBox(mainwin, "Done!",
			fmt.Sprintf("%d matching PN(s) found from %d vendor file(s).\n"+
				"%d price cell(s) updated.\n"+
				"%d lead time cell(s) updated.",
				len(pnFoundMap), len(vendorEntries), priceUpdateCounter, leadTimeUpdatecounter))
	})
}

func main() {
	ui.Main(setupUI)
}
