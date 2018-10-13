package main

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"github.com/andlabs/ui"
	"github.com/tealeg/xlsx"
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
	r := csv.NewReader(bufio.NewReader(f))
	file := xlsx.NewFile()
	sheet, err := file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf(err.Error())
	}

	for {
		record, err := r.Read()
		// Stop at EOF.
		if err == io.EOF {
			break
		}
		row := sheet.AddRow()

		for _, value := range record {
			cell := row.AddCell()
			cell.Value = value
		}
	}

	err = file.Save(target)
	if err != nil {
		fmt.Printf(err.Error())
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
	//
	//tab.Append("Numbers and Lists", makeNumbersPage())
	//tab.SetMargined(1, true)
	//
	//tab.Append("Data Choosers", makeDataChoosersPage())
	//tab.SetMargined(2, true)

	mainwin.Show()
}

func main() {
	ui.Main(setupUI)
}
