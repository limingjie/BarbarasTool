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

// GUI - The GUI of the tool
type GUI struct {
	window               *ui.Window
	vendorFileEntriesBox *ui.Box
	vendorFileEntries    []*ui.Entry // Keep track of vendor files entries for removal
	bidFileEntry         *ui.Entry
}

// Open a xlsx file
// - required, a file must be selected.
func (gui *GUI) openExcelFile(required bool) (filename string) {
	for true {
		filename = ui.OpenFile(gui.window)

		// The open file dialog is cancelled
		if len(filename) == 0 {
			if required {
				ui.MsgBoxError(gui.window, "Error", "A file is required. Please try again.")
				continue
			} else {
				break
			}
		}

		// Not an xlsx file
		if !strings.HasSuffix(filename, ".xlsx") {
			ui.MsgBoxError(gui.window, "Error", "Only support xlsx format. Please try again.")
			continue
		}

		break
	}

	return
}

// Remove the vendor file entry and destroy the UI
func (gui *GUI) onRemoveButtonClick(entryToRemove *ui.Entry, entryBox *ui.Box) {
	for i, e := range gui.vendorFileEntries {
		if e == entryToRemove {
			// Destroy the file entry UI
			gui.vendorFileEntriesBox.Delete(i)
			// Destroying box cascades to all box children
			entryBox.Destroy()
			// Remove the entry from the tracking slice
			gui.vendorFileEntries = append(gui.vendorFileEntries[:i], gui.vendorFileEntries[i+1:]...)
			break
		}
	}
}

func (gui *GUI) onProcessButtonClick() {
	// Create data maps by reading all vendor file sheet(s)
	// - PN        - Col C - "型号（P/N)"
	// - Price     - Col H - "合计（不含税）"
	// - Lead time - Col L - "货期"
	priceMap := make(map[string]string)
	leadTimeMap := make(map[string]string)
	for _, entry := range gui.vendorFileEntries {
		_, err := appendExcelDataToMap(&priceMap, entry.Text(), 'C'-'A', 'H'-'A')
		if err != nil {
			ui.MsgBoxError(gui.window, "Error", err.Error())
		}
		_, err = appendExcelDataToMap(&leadTimeMap, entry.Text(), 'C'-'A', 'L'-'A')
		if err != nil {
			ui.MsgBoxError(gui.window, "Error", err.Error())
		}
	}

	// Convert some Chinese words to English
	for k, v := range leadTimeMap {
		v = strings.ReplaceAll(v, "周", " wks")
		v = strings.ReplaceAll(v, "现货", "In stock")
		leadTimeMap[k] = v
	}

	// Stop if both price map and lead time map are empty
	if len(priceMap) == 0 && len(leadTimeMap) == 0 {
		ui.MsgBoxError(gui.window, "Error", "No price or lead time data found from vendor file(s)!")
		return
	}

	// Open bid file
	bidFile, err := excelize.OpenFile(gui.bidFileEntry.Text())
	if err != nil {
		ui.MsgBoxError(gui.window, "Error", err.Error())
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
			ui.MsgBoxError(gui.window, "Error", err.Error())
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
				if _, err := strconv.ParseFloat(price, 64); err == nil {
					if length < 11 || price != rows[i]['K'-'A'] {
						if cellName, err := excelize.CoordinatesToCellName('K'-'A'+1, i+1); err == nil {
							bidFile.SetCellDefault(sheetName, cellName, price)
							priceUpdateCounter++
						}
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
			ui.MsgBoxError(gui.window, "Error", "Error save bid file.\n"+err.Error())
			return
		}
	}

	// Done
	ui.MsgBox(gui.window, "Done!",
		fmt.Sprintf("%d matching PN(s) found from %d vendor file(s).\n"+
			"%d price cell(s) updated.\n"+
			"%d lead time cell(s) updated.",
			len(pnFoundMap), len(gui.vendorFileEntries), priceUpdateCounter, leadTimeUpdatecounter))
}

// Add a new vendor file
func (gui *GUI) addVendorFile() (added bool) {
	// Open file, the first vendor file is required.
	filename := gui.openExcelFile(len(gui.vendorFileEntries) == 0)

	if len(filename) != 0 {
		// Add a new entry to UI
		entry := ui.NewEntry()
		entry.SetReadOnly(true)
		entry.SetText(filename)

		// Remove button
		button := ui.NewButton(" - ")

		// | <- entry -> | button |
		entryBox := ui.NewHorizontalBox()
		entryBox.SetPadded(true)
		entryBox.Append(entry, true)
		entryBox.Append(button, false)

		// Remove the vendor file entry on remove button click
		button.OnClicked(func(*ui.Button) {
			gui.onRemoveButtonClick(entry, entryBox)
		})

		// Append and track the new vendor entry
		gui.vendorFileEntriesBox.Append(entryBox, false)
		gui.vendorFileEntries = append(gui.vendorFileEntries, entry)

		return true
	}

	return false
}

func newGUI() {
	gui := &GUI{
		window:               ui.NewWindow("Barbara's Tool", 700, 100, false),
		vendorFileEntriesBox: ui.NewVerticalBox(),
		vendorFileEntries:    make([]*ui.Entry, 0),
		bidFileEntry:         ui.NewEntry(),
	}

	gui.window.OnClosing(func(*ui.Window) bool {
		ui.Quit()
		return true
	})

	ui.OnShouldQuit(func() bool {
		gui.window.Destroy()
		return true
	})

	// Add vendor file button
	addButton := ui.NewButton("Add Vendor File")
	addButton.OnClicked(func(*ui.Button) {
		gui.addVendorFile()
	})

	// Vendor files box
	gui.vendorFileEntriesBox.SetPadded(true)

	// Bid file entry
	gui.bidFileEntry.SetReadOnly(true)

	// Reset bid file button
	resetButton := ui.NewButton(" R ")
	resetButton.OnClicked(func(*ui.Button) {
		filename := gui.openExcelFile(false)
		if len(filename) != 0 {
			gui.bidFileEntry.SetText(filename)
		}
	})

	// Process button
	processButton := ui.NewButton("Process")
	processButton.OnClicked(func(*ui.Button) {
		gui.onProcessButtonClick()
	})

	// Main layout
	box := ui.NewVerticalBox()
	box.SetPadded(true)
	box.Append(ui.NewLabel("Vendor file - Search Price (H) & Lead time (L) base on P/N (C)"), false)
	box.Append(gui.vendorFileEntriesBox, false)
	// Vertical stretch
	box.Append(ui.NewVerticalBox(), true)
	// | <- Stretch -> | addButton |
	box.Append(func() ui.Control {
		hBox := ui.NewHorizontalBox()
		hBox.SetPadded(true)
		hBox.Append(ui.NewLabel(""), true)
		hBox.Append(addButton, false)
		return hBox
	}(), false)
	box.Append(ui.NewHorizontalSeparator(), false)
	box.Append(ui.NewLabel("Bid file - Fill in Price (K) & Lead time (P) base on P/N (F)"), false)
	// | <- bidFileEntry -> | resetButton |
	box.Append(func() ui.Control {
		hBox := ui.NewHorizontalBox()
		hBox.SetPadded(true)
		hBox.Append(gui.bidFileEntry, true)
		hBox.Append(resetButton, false)
		return hBox
	}(), false)
	box.Append(ui.NewHorizontalSeparator(), false)
	// | Cheers Barbara! | <- Stretch -> | processButton |
	box.Append(func() ui.Control {
		hBox := ui.NewHorizontalBox()
		hBox.SetPadded(true)
		hBox.Append(ui.NewLabel("Cheers Barbara!"), false)
		hBox.Append(ui.NewLabel(""), true)
		hBox.Append(processButton, false)
		return hBox
	}(), false)

	// Main window
	gui.window.SetChild(box)
	gui.window.SetMargined(true)
	gui.window.Show()

	// Open vendor files
	ui.MsgBox(gui.window, "Message", "Open the vendor files.")
	for gui.addVendorFile() {
		// Until cancelled
	}

	// Open bid files
	ui.MsgBox(gui.window, "Message", "Open the bid file.")
	gui.bidFileEntry.SetText(gui.openExcelFile(true))
}

func appendExcelDataToMap(dataMap *map[string]string, filename string, keyCol, valueCol int) (*map[string]string, error) {
	file, err := excelize.OpenFile(filename)
	if err != nil {
		return dataMap, err
	}

	for _, sheetName := range file.GetSheetMap() {
		rows, err := file.GetRows(sheetName)
		if err != nil {
			continue
		}
		for i := 0; i < len(rows); i++ {
			// Make sure key & value columns are exist
			if len(rows[i]) < keyCol+1 || len(rows[i]) < valueCol+1 {
				continue
			}

			key := strings.ToLower(strings.TrimSpace(rows[i][keyCol]))
			value := strings.TrimSpace(rows[i][valueCol])

			if len(key) != 0 && len(value) != 0 {
				(*dataMap)[key] = value
			}
		}
	}

	return dataMap, nil
}

func main() {
	ui.Main(newGUI)
}
