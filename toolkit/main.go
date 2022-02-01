package main

import (
	"bufio"
	"encoding/csv"
	"fmt"
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
				fmt.Println(err)
			}
			err = excel.SetCellValue(sheetName, axis, value)
			if err != nil {
				return
			}
		}
		lineNum += 1
	}
	if err := excel.SaveAs(target); err != nil {
		fmt.Println(err)
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
	box := ui.NewVerticalBox()
	box.SetPadded(true)
	box.Append(ui.NewLabel("Drop file path below"), false)
	box.Append(input, false)
	box.Append(button, false)
	box.Append(resultMsg, false)
	return box
}

func setupUI() {
	mainwin = ui.NewWindow("ToolKit", 640, 480, false)
	mainwin.OnClosing(func(*ui.Window) bool {
		ui.Quit()
		return true
	})
	ui.OnShouldQuit(func() bool {
		mainwin.Destroy()
		return true
	})

	tab := ui.NewTab()
	mainwin.SetChild(tab)
	mainwin.SetMargined(true)

	tab.Append("csv->xlsx", makeConvertPage())
	tab.SetMargined(0, true)

	mainwin.Show()
}

func main() {
	ui.Main(setupUI)
}
