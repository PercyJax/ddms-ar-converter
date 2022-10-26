package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/gosimple/slug"
	"github.com/sqweek/dialog"
	"github.com/xuri/excelize/v2"
)

var workingDir string
var mainSheet = "Accounts"

func getReportPath() string {
	path, err := os.Getwd()
	if err != nil {
		dialog.Message(fmt.Sprintf("%s\n", err)).Title("Error").Error()
	}
	workingDir = path
	f := &dialog.FileBuilder{
		Dlg: dialog.Dlg{
			Title: "Select an A/R Report file...",
		},
		StartDir: path,
		Filters: []dialog.FileFilter{
			{
				Desc:       "AR Report File",
				Extensions: []string{"txt"},
			},
		},
	}
	filename, err := f.Load()
	if err == dialog.ErrCancelled {
		os.Exit(0)
	}
	return filename
}

func convertToExcel(txtFile string) {
	// Open .txt file
	t, err := os.Open(txtFile)
	if err != nil {
		dialog.Message(fmt.Sprintf("%s\n", err)).Title("Error").Error()
		os.Exit(1)
	}
	defer t.Close()
	txt := bufio.NewScanner(t)

	var company, address, city, state, zip string
	var report_title, report_subtitle string
	var report_date, report_time string

	var header string
	header_lines := 6

	// Read first line to grab company name, title, and date
	if txt.Scan() {
		line := (txt.Text())
		header = line
		company = strings.TrimSpace(line[0:58])
		report_title = strings.TrimSpace(line[58:121])
		report_date = strings.TrimSpace(line[121:])
	} else {
		dialog.Message("File empty").Title("Error").Error()
		os.Exit(1)
	}

	// Read second line to address, title line 2, and timestamp
	if txt.Scan() {
		line := (txt.Text())
		address = strings.TrimSpace(line[0:59])
		report_title = report_title + " " + strings.TrimSpace(line[59:121])
		report_time = strings.TrimSpace(line[121:])
	} else {
		dialog.Message("Report is corrupted").Title("Error").Error()
		os.Exit(1)
	}

	// Read third line for address line 2, title line 3
	if txt.Scan() {
		line := (txt.Text())
		city = strings.TrimSpace(line[0:21])
		state = strings.TrimSpace(line[21:23])
		zip = strings.TrimSpace(line[23:60])
		report_subtitle = strings.TrimSpace(line[60:121])
	} else {
		dialog.Message("Report is corrupted").Title("Error").Error()
		os.Exit(1)
	}

	// Read fourth line for aged by date
	if txt.Scan() {
		line := (txt.Text())
		report_title = report_title + " " + strings.TrimSpace(line)
	} else {
		dialog.Message("Report is corrupted").Title("Error").Error()
		os.Exit(1)
	}

	// Read fifth line for spacing info?
	if !txt.Scan() {
		dialog.Message("Report is corrupted").Title("Error").Error()
		os.Exit(1)
	}
	// Skip sixth line
	if !txt.Scan() {
		dialog.Message("Report is corrupted").Title("Error").Error()
		os.Exit(1)
	}

	// fmt.Println(company)
	// fmt.Println(report_date)
	// fmt.Println(report_time)
	// fmt.Println(address)
	// fmt.Println(city+", ", state, zip)
	// fmt.Println(report_title)

	// Add excel sheet header
	x := excelize.NewFile()
	x.SetSheetName("Sheet1", mainSheet)
	x.SetCellValue(mainSheet, "B1", company)
	x.MergeCell(mainSheet, "B1", "G1")
	x.SetCellValue(mainSheet, "B2", address)
	x.MergeCell(mainSheet, "B2", "G2")
	x.SetCellValue(mainSheet, "B3", city+", "+state+" "+zip)
	x.MergeCell(mainSheet, "B3", "G3")

	x.SetCellValue(mainSheet, "I1", report_title)
	x.MergeCell(mainSheet, "I1", "N1")
	x.SetCellValue(mainSheet, "I2", report_subtitle)
	x.MergeCell(mainSheet, "I2", "N2")
	x.SetCellValue(mainSheet, "I3", "Report Generated: "+report_date+" at "+report_time)
	x.MergeCell(mainSheet, "I3", "N3")

	// Header Row
	row := 5

	// Create columns: Invoice #, Name, Address 1, Address 2, City, State, Zip, Current, Over 30, Over 60, Over 90, Discount, Balance
	x.SetCellValue(mainSheet, "A"+strconv.Itoa(row), "INVOICE #")
	x.SetCellValue(mainSheet, "B"+strconv.Itoa(row), "CUSTOMER")
	x.SetCellValue(mainSheet, "C"+strconv.Itoa(row), "ADDRESS LINE 1")
	x.SetCellValue(mainSheet, "D"+strconv.Itoa(row), "ADDRESS LINE 2")
	x.SetCellValue(mainSheet, "E"+strconv.Itoa(row), "CITY")
	x.SetCellValue(mainSheet, "F"+strconv.Itoa(row), "STATE")
	x.SetCellValue(mainSheet, "G"+strconv.Itoa(row), "ZIP")
	x.SetCellValue(mainSheet, "H"+strconv.Itoa(row), "SALESMAN")
	x.SetCellValue(mainSheet, "I"+strconv.Itoa(row), "CURRENT")
	x.SetCellValue(mainSheet, "J"+strconv.Itoa(row), "OVER 30")
	x.SetCellValue(mainSheet, "K"+strconv.Itoa(row), "OVER 60")
	x.SetCellValue(mainSheet, "L"+strconv.Itoa(row), "OVER 90")
	x.SetCellValue(mainSheet, "M"+strconv.Itoa(row), "DISCOUNT")
	x.SetCellValue(mainSheet, "N"+strconv.Itoa(row), "BALANCE")

	// Set visibility
	x.SetColVisible(mainSheet, "C:D", false)
	x.SetColVisible(mainSheet, "H", false)

	// Set column widths
	x.SetColWidth(mainSheet, "A", "A", 15) //10
	x.SetColWidth(mainSheet, "B", "B", 40) //30
	x.SetColWidth(mainSheet, "C", "D", 30) //25
	x.SetColWidth(mainSheet, "E", "E", 19)
	x.SetColWidth(mainSheet, "F", "F", 10) //6
	x.SetColWidth(mainSheet, "G", "G", 8)  //5
	x.SetColWidth(mainSheet, "H", "H", 66)
	x.SetColWidth(mainSheet, "I", "N", 15) //11

	// Read every 3 lines, convert to excel, check if match to line 1 (new page), or blank line
	complete := false
	data_first_row := row + 1
	data_last_row := data_first_row
	for txt.Scan() {
		line := txt.Text()
		// fmt.Println(line)

		if strings.TrimSpace(line) == strings.TrimSpace(header) {
			// fmt.Println("Header Match")
			// Skip headers if match
			for i := 0; i < header_lines-1; i++ {
				txt.Scan()
			}
		} else if strings.Contains(line, "TOTAL OF ALL A/R PRINTED") {
			// End loop if "TOTAL OF ALL A/R PRINTED"
			// Capture totals
			complete = true
			break
		} else if line == "" {
			dialog.Message("Report is corrupted").Title("Error").Error()
			os.Exit(1)
		} else {
			row++
			var invoice int
			var customer, address1, address2, city, state, zip, salesman string
			var current, over30, over60, over90, discount, balance float64

			invoice, err = strconv.Atoi(strings.TrimSpace(line[0:10]))
			if err != nil {
				dialog.Message(fmt.Sprintf("%s\n", err)).Title("Error").Error()
				os.Exit(1)
			}
			customer = strings.TrimSpace(line[11:41])
			address1 = strings.TrimSpace(line[42:67])
			address2 = strings.TrimSpace(line[68:93])
			city = strings.TrimSpace(line[94:113])
			state = strings.TrimSpace(line[114:116])
			zip = strings.TrimSpace(line[117:])

			if !txt.Scan() {
				dialog.Message("Report is corrupted").Title("Error").Error()
				os.Exit(1)
				// break
			} else {
				line = txt.Text()
				salesman = strings.TrimSpace(line)
			}

			if !txt.Scan() {
				dialog.Message("Report is corrupted").Title("Error").Error()
				os.Exit(1)
				// break
			} else {
				line = txt.Text()
				// fmt.Println(line)

				var val string

				val = strings.TrimSpace(line[61:72])
				if val == "" {
					current = 0
				} else {
					current, err = strconv.ParseFloat(val, 64)
					if err != nil {
						dialog.Message(fmt.Sprintf("%s\n", err)).Title("Error").Error()
						os.Exit(1)
					}
				}

				val = strings.TrimSpace(line[73:84])
				if val == "" {
					over30 = 0
				} else {
					over30, err = strconv.ParseFloat(val, 64)
					if err != nil {
						dialog.Message(fmt.Sprintf("%s\n", err)).Title("Error").Error()
						os.Exit(1)
					}
				}

				val = strings.TrimSpace(line[85:96])
				if val == "" {
					over60 = 0
				} else {
					over60, err = strconv.ParseFloat(val, 64)
					if err != nil {
						dialog.Message(fmt.Sprintf("%s\n", err)).Title("Error").Error()
						os.Exit(1)
					}
				}

				val = strings.TrimSpace(line[97:108])
				if val == "" {
					over90 = 0
				} else {
					over90, err = strconv.ParseFloat(val, 64)
					if err != nil {
						dialog.Message(fmt.Sprintf("%s\n", err)).Title("Error").Error()
						os.Exit(1)
					}
				}

				val = strings.TrimSpace(line[109:117])
				if val == "" {
					discount = 0
				} else {
					discount, err = strconv.ParseFloat(val, 64)
					if err != nil {
						dialog.Message(fmt.Sprintf("%s\n", err)).Title("Error").Error()
						os.Exit(1)
					}
				}

				val = strings.TrimSpace(line[118:])
				if val == "" {
					balance = 0
				} else {
					balance, err = strconv.ParseFloat(val, 64)
					if err != nil {
						dialog.Message(fmt.Sprintf("%s\n", err)).Title("Error").Error()
						os.Exit(1)
					}
				}
			}

			x.SetCellValue(mainSheet, "A"+strconv.Itoa(row), invoice)
			x.SetCellValue(mainSheet, "B"+strconv.Itoa(row), customer)
			x.SetCellValue(mainSheet, "C"+strconv.Itoa(row), address1)
			x.SetCellValue(mainSheet, "D"+strconv.Itoa(row), address2)
			x.SetCellValue(mainSheet, "E"+strconv.Itoa(row), city)
			x.SetCellValue(mainSheet, "F"+strconv.Itoa(row), state)
			x.SetCellValue(mainSheet, "G"+strconv.Itoa(row), zip)
			x.SetCellValue(mainSheet, "H"+strconv.Itoa(row), salesman)
			x.SetCellValue(mainSheet, "I"+strconv.Itoa(row), current)
			x.SetCellValue(mainSheet, "J"+strconv.Itoa(row), over30)
			x.SetCellValue(mainSheet, "K"+strconv.Itoa(row), over60)
			x.SetCellValue(mainSheet, "L"+strconv.Itoa(row), over90)
			x.SetCellValue(mainSheet, "M"+strconv.Itoa(row), discount)
			x.SetCellValue(mainSheet, "N"+strconv.Itoa(row), balance)
			data_last_row = row
		}
	}
	// fmt.Println(complete)

	// Process footer/calculations
	row = row + 2
	x.SetCellValue(mainSheet, "B"+strconv.Itoa(row), "TOTAL OF ALL A/R PRINTED")
	x.MergeCell(mainSheet, "B"+strconv.Itoa(row), "H"+strconv.Itoa(row))
	x.SetCellFormula(mainSheet, "I"+strconv.Itoa(row), "=SUBTOTAL(9, I"+strconv.Itoa(data_first_row)+":I"+strconv.Itoa(data_last_row)+")")
	x.SetCellFormula(mainSheet, "J"+strconv.Itoa(row), "=SUBTOTAL(9, J"+strconv.Itoa(data_first_row)+":J"+strconv.Itoa(data_last_row)+")")
	x.SetCellFormula(mainSheet, "K"+strconv.Itoa(row), "=SUBTOTAL(9, K"+strconv.Itoa(data_first_row)+":K"+strconv.Itoa(data_last_row)+")")
	x.SetCellFormula(mainSheet, "L"+strconv.Itoa(row), "=SUBTOTAL(9, L"+strconv.Itoa(data_first_row)+":L"+strconv.Itoa(data_last_row)+")")
	x.SetCellFormula(mainSheet, "M"+strconv.Itoa(row), "=SUBTOTAL(9, M"+strconv.Itoa(data_first_row)+":M"+strconv.Itoa(data_last_row)+")")
	x.SetCellFormula(mainSheet, "N"+strconv.Itoa(row), "=SUBTOTAL(9, N"+strconv.Itoa(data_first_row)+":N"+strconv.Itoa(data_last_row)+")")

	// Set styles
	var currency_style int
	if currency_style, err = x.NewStyle(&excelize.Style{
		NumFmt:        40,
		DecimalPlaces: 2,
	}); err != nil {
		dialog.Message(fmt.Sprintf("Error creating style: %s\n", err)).Title("Error").Error()
		os.Exit(1)
	}
	if err = x.SetColStyle(mainSheet, "I:N", currency_style); err != nil {
		dialog.Message(fmt.Sprintf("Error setting style: %s\n", err)).Title("Error").Error()
		os.Exit(1)
	}

	if err := x.SetPanes(mainSheet, `{
		"freeze": true,
		"split": false,
		"x_split": 1,
		"y_split": `+strconv.Itoa(data_first_row-1)+`,
		"top_left_cell": "B`+strconv.Itoa(data_first_row)+`",
		"active_pane": "bottomRight",
		"panes": [
		{
			"sqref": "B`+strconv.Itoa(data_first_row)+`",
			"active_cell": "B`+strconv.Itoa(data_first_row)+`",
			"pane": "bottomRight"
		}]
	}`); err != nil {
		dialog.Message(fmt.Sprintf("Error freezing panes: %s\n", err)).Title("Error").Error()
		os.Exit(1)
	}

	var header_style int
	if header_style, err = x.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	}); err != nil {
		dialog.Message(fmt.Sprintf("Error creating style: %s\n", err)).Title("Error").Error()
		os.Exit(1)
	}
	if err = x.SetRowStyle(mainSheet, data_first_row-1, data_first_row-1, header_style); err != nil {
		dialog.Message(fmt.Sprintf("Error setting style: %s\n", err)).Title("Error").Error()
		os.Exit(1)
	}

	var right_justified_style int
	var centered_style int

	var current_style int

	if current_style, err = x.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "right",
		},
	}); err != nil {
		dialog.Message(fmt.Sprintf("Error creating style: %s\n", err)).Title("Error").Error()
		os.Exit(1)
	}
	right_justified_style = current_style

	if current_style, err = x.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	}); err != nil {
		dialog.Message(fmt.Sprintf("Error creating style: %s\n", err)).Title("Error").Error()
		os.Exit(1)
	}
	centered_style = current_style

	current_style = right_justified_style
	if err = x.SetCellStyle(mainSheet, "B"+strconv.Itoa(data_last_row+2), "B"+strconv.Itoa(data_last_row+2), current_style); err != nil {
		dialog.Message(fmt.Sprintf("Error setting style: %s\n", err)).Title("Error").Error()
		os.Exit(1)
	}
	if err = x.SetCellStyle(mainSheet, "I1", "N3", current_style); err != nil {
		dialog.Message(fmt.Sprintf("Error setting style: %s\n", err)).Title("Error").Error()
		os.Exit(1)
	}

	current_style = centered_style
	if err = x.SetCellStyle(mainSheet, "B1", "B"+strconv.Itoa(data_first_row-2), current_style); err != nil {
		dialog.Message(fmt.Sprintf("Error setting style: %s\n", err)).Title("Error").Error()
		os.Exit(1)
	}

	// Add filter
	x.AutoFilter(mainSheet, "A"+strconv.Itoa(data_first_row-1), "N"+strconv.Itoa(data_first_row-1), "")

	// Save file
	if complete {
		f := &dialog.FileBuilder{
			Dlg: dialog.Dlg{
				Title: "Save A/R Report (Excel Sheet)...",
			},
			StartDir:  workingDir,
			StartFile: slug.Make(company+"_"+report_date+"_"+report_time) + ".xlsx",
			Filters: []dialog.FileFilter{
				{
					Desc:       "AR Report File",
					Extensions: []string{"xlsx"},
				},
			},
		}
		filename, _ := f.Save()
		if err = x.SaveAs(filename); err != nil {
			dialog.Message(fmt.Sprintf("%s\n", err)).Title("Error").Error()
			os.Exit(1)
		}
	}
}

func main() {
	convertToExcel(getReportPath())
}
