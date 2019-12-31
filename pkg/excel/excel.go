package excel

import (
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// IndexColumns - Index key and value columns of given file.
func IndexColumns(data *map[string]string, filename string, keyCol, valueCol int) (*map[string]string, error) {
	file, err := excelize.OpenFile(filename)
	if err != nil {
		return data, err
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
				(*data)[key] = value
			}
		}
	}

	return data, nil
}

// UpdateColumnByIndex - Update the file using index data
func UpdateColumnByIndex(data *map[string]string,
	filename string, keyCol, valueCol int, formattedValue bool,
) (found, updated int, err error) {
	// Open file
	file, err := excelize.OpenFile(filename)
	if err != nil {
		return
	}

	// Update file
	foundMap := make(map[string]bool)
	for _, sheetName := range file.GetSheetMap() {
		rows, err := file.GetRows(sheetName)
		if err != nil {
			// Omit and continue with next sheet
			continue
		}

		for i := 0; i < len(rows); i++ {
			// Make sure key column exists
			length := len(rows[i])
			if len(rows[i]) < keyCol+1 {
				continue
			}

			// Get key
			key := strings.ToLower(strings.TrimSpace(rows[i][keyCol]))
			if len(key) == 0 {
				continue
			}

			// Update value
			if value, ok := (*data)[key]; ok {
				foundMap[key] = true
				if length < valueCol || value != rows[i][valueCol] {
					if cellName, err := excelize.CoordinatesToCellName(valueCol+1, i+1); err == nil {
						if formattedValue {
							if _, err := strconv.ParseFloat(value, 64); err == nil {
								file.SetCellDefault(sheetName, cellName, value)
								updated++
							}
						} else {
							file.SetCellStr(sheetName, cellName, value)
							updated++
						}
					}
				}
			}
		}
	}

	// Save file if updated
	if updated != 0 {
		return len(foundMap), updated, file.Save()
	}

	return len(foundMap), 0, nil
}

// Backup - Take a backup
func Backup(filename string) error {
	file, err := excelize.OpenFile(filename)
	if err != nil {
		return err
	}
	return file.SaveAs("./backup." + time.Now().Format("20060102.150405.999999") + ".xlsx")
}
