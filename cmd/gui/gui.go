package gui

import (
	"fmt"
	"strings"

	"github.com/andlabs/ui"
	"github.com/limingjie/BarbarasTool/pkg/excel"
)

// GUI - The GUI of the tool
type GUI struct {
	window               *ui.Window
	vendorFileEntriesBox *ui.Box
	vendorFileEntries    []*ui.Entry // Keep track of vendor files entries for removal
	bidFileEntry         *ui.Entry
}

// NewGUI - Create GUI
func NewGUI() {
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
	// | ---------------------------------------------------------------|
	// | Vendor file - Search Price (H) & Lead time (L) base on P/N (C) |
	// |                    <- Vendor File 0 ->                   | [-] |
	// |                    <- Vendor File 1 ->                   | [-] |
	// |                                            | [Add Vendor File] |
	// | ---------------------------------------------------------------|
	// | Bid file - Fill in Price (K) & Lead time (P) base on P/N (F)   |
	// |                    <-   Bid File    ->                   | [R] |
	// | ---------------------------------------------------------------|
	// | Cheers Barbara! |                                  | [Process] |
	// | ---------------------------------------------------------------|
	box := ui.NewVerticalBox()
	box.SetPadded(true)
	box.Append(ui.NewLabel("Vendor file - Search Price (H) & Lead time (L) base on P/N (C)"), false)
	box.Append(gui.vendorFileEntriesBox, false)
	box.Append(ui.NewVerticalBox(), true)
	box.Append(func() ui.Control {
		hBox := ui.NewHorizontalBox()
		hBox.SetPadded(true)
		hBox.Append(ui.NewLabel(""), true)
		hBox.Append(addButton, false)
		return hBox
	}(), false)
	box.Append(ui.NewHorizontalSeparator(), false)
	box.Append(ui.NewLabel("Bid file - Fill in Price (K) & Lead time (P) base on P/N (F)"), false)
	box.Append(func() ui.Control {
		hBox := ui.NewHorizontalBox()
		hBox.SetPadded(true)
		hBox.Append(gui.bidFileEntry, true)
		hBox.Append(resetButton, false)
		return hBox
	}(), false)
	box.Append(ui.NewHorizontalSeparator(), false)
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

// Open an xlsx file
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
		if !strings.HasSuffix(strings.ToLower(filename), ".xlsx") {
			ui.MsgBoxError(gui.window, "Error", "Only support xlsx format. Please try again.")
			continue
		}

		break
	}

	return
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
		_, err := excel.IndexColumns(&priceMap, entry.Text(), 'C'-'A', 'H'-'A')
		if err != nil {
			ui.MsgBoxError(gui.window, "Error", err.Error())
		}
		_, err = excel.IndexColumns(&leadTimeMap, entry.Text(), 'C'-'A', 'L'-'A')
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

	// Backup bid file
	bidFilename := gui.bidFileEntry.Text()
	excel.Backup(bidFilename)

	// Update bid file using maps
	// - PN        - Col F - International Part No.
	// - Price     - Col K - Unit Price CNY
	// - Lead time - Col P - leadtime
	summary := fmt.Sprintf("In %d vendor file(s).", len(gui.vendorFileEntries))

	found, updated, err := excel.UpdateColumnByIndex(&priceMap, bidFilename, 'F'-'A', 'K'-'A', true)
	if err != nil {
		ui.MsgBoxError(gui.window, "Error", err.Error())
		return
	}
	summary += fmt.Sprintf("\n- %d matching PN(s) found, %d price(s) updated.", found, updated)

	found, updated, err = excel.UpdateColumnByIndex(&leadTimeMap, bidFilename, 'F'-'A', 'P'-'A', false)
	if err != nil {
		ui.MsgBoxError(gui.window, "Error", err.Error())
		return
	}
	summary += fmt.Sprintf("\n- %d matching PN(s) found, %d lead time updated.", found, updated)

	// Done
	ui.MsgBox(gui.window, "Done!", summary)
}
