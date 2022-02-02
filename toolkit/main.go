package main

import (
	"bufio"
	"encoding/csv"
	"github.com/andlabs/ui"
	"github.com/xuri/excelize/v2"
	"io"
	"os"
	"strings"
)

var mainwin *ui.Window

func convert(source, target string) {
	f, err := os.Open(source)
	if err != nil {
		return
	}
	defer f.Close()
	r := csv.NewReader(bufio.NewReader(f))
	excel := excelize.NewFile()
	sheetName := "Sheet1"
	excel.SetActiveSheet(excel.NewSheet(sheetName))
	lineNum := 1
	for {
		record_line, err := r.Read()
		// Stop at EOF.
		if err == io.EOF {
			break
		}
		for columnNum, value := range record_line {
			axis, err := excelize.CoordinatesToCellName(columnNum+1, lineNum)
			if err != nil {
				return
			}
			err = excel.SetCellValue(sheetName, axis, value)
			if err != nil {
				return
			}
		}
		lineNum += 1
	}
	if err := excel.SaveAs(target); err != nil {
		return
	}
}

func makeConvertPage() ui.Control {
	input := ui.NewMultilineEntry()
	button := ui.NewButton("Convert")
	resultMsg := ui.NewLabel("")

	button.OnClicked(func(*ui.Button) {
		source := input.Text()
		target := strings.Replace(source, "csv", "xlsx", 1)

		input.SetText("")

		resultMsg.SetText("convert to:" + target)
		convert(source, target)
		resultMsg.SetText("convert to:" + target + ", finish!")
	})
	openFileButton := ui.NewButton("Open File")
	openFileButton.OnClicked(func(*ui.Button) {
		filename := ui.OpenFile(mainwin)
		if filename != "" {
			if strings.HasSuffix(filename, ".csv") {
				input.SetText(filename)
			} else {
				ui.MsgBoxError(mainwin,
					"ERROR",
					"please select file end with .csv")
			}
		}
	})
	box := ui.NewVerticalBox()
	box.SetPadded(true)

	hbox := ui.NewHorizontalBox()
	hbox.SetPadded(true)
	hbox.Append(ui.NewLabel("Drop file path below or click Open File button"), true)
	hbox.Append(openFileButton, false)

	box.Append(hbox, false)
	box.Append(input, false)
	box.Append(button, false)
	box.Append(resultMsg, false)
	return box
}

func setupUI() {
	mainwin = ui.NewWindow("ToolKit", 500, 400, false)
	mainwin.OnClosing(func(*ui.Window) bool {
		ui.Quit()
		return true
	})
	ui.OnShouldQuit(func() bool {
		mainwin.Destroy()
		return true
	})
	mainwin.SetChild(makeConvertPage())
	mainwin.SetMargined(true)
	mainwin.Show()
}

func main() {
	ui.Main(setupUI)
}
