package main

import (
	"fmt"
	"image/color"
	"math"
	"os"
	"strconv"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
	"gonum.org/v1/plot"
	"gonum.org/v1/plot/plotter"
	"gonum.org/v1/plot/vg"
)

var (
	scale   float64
	Tbaz    float64
	Tzadan  float64
	Zs      float64
	zLinear float64
	NLinear float64
	zExp    float64
	NExp    float64
	image   *canvas.Image
	resultLabel *widget.Label
)

func saveReport() {
	f := excelize.NewFile()
	sheet := f.GetSheetName(0)

	f.SetCellValue(sheet, "A1", "Параметр")
	f.SetCellValue(sheet, "B1", "Значение")
	params := map[string]float64{
		"Масштаб":              scale,
		"Базовый срок службы":  Tbaz,
		"Нахождение затрат":    Tzadan,
		"Распределение затрат": Zs,
		"z (линейная)":         zLinear,
		"N (линейная)":         NLinear,
		"z (экспонент.)":       zExp,
		"N (экспонент.)":       NExp,
		"Z на подсистему":      Zs / 4,
	}	

	row := 2
	for k, v := range params {
		f.SetCellValue(sheet, fmt.Sprintf("A%d", row), k)
		f.SetCellValue(sheet, fmt.Sprintf("B%d", row), v)
		row++
	}

	if _, err := os.Stat("plot.png"); err == nil {
		opt := &excelize.GraphicOptions{ScaleX: 0.7, ScaleY: 0.7}
		if err := f.AddPicture(sheet, "D2", "plot.png", opt); err != nil {
			fmt.Println("Ошибка вставки картинки:", err)
		}
	}

	filename := fmt.Sprintf("отчёт_%s.xlsx", time.Now().Format("20060102_150405"))
	if err := f.SaveAs(filename); err != nil {
		fmt.Println("Ошибка сохранения:", err)
	} else {
		fmt.Println("Отчёт сохранён как", filename)
	}
}

func main() {
	myApp := app.New()
	myWindow := myApp.NewWindow("Распределение затрат")

	scaleEntry := widget.NewEntry()
	scaleEntry.SetText("200")

	TbazEntry := widget.NewEntry()
	TbazEntry.SetText("8")

	TzadanEntry := widget.NewEntry()
	TzadanEntry.SetText("2")

	ZsEntry := widget.NewEntry()
	ZsEntry.SetText("19")

	resultLabel = widget.NewLabel("Здесь появятся результаты")
	image = canvas.NewImageFromFile("")
	image.FillMode = canvas.ImageFillOriginal

	calcButton := widget.NewButton("Рассчитать", func() {
		scale, _ = strconv.ParseFloat(scaleEntry.Text, 64)
		Tbaz, _ = strconv.ParseFloat(TbazEntry.Text, 64)
		Tzadan, _ = strconv.ParseFloat(TzadanEntry.Text, 64)
		Zs, _ = strconv.ParseFloat(ZsEntry.Text, 64)
		
		if Tzadan >= Tbaz {
			resultLabel.SetText("Ошибка: срок службы подсистемы должен быть меньше базового")
			return
		}

		k := Zs / Tbaz
		zLinear = Tzadan / k
		NLinear = Zs / zLinear
		zExp = -(1 / k) * math.Log(1-(Tzadan/Tbaz))
		NExp = Zs / zExp

		res := fmt.Sprintf("Масштаб: %.0f\nЛинейная модель: z=%.2f, N=%.1f\nЭкспоненциальная модель: z=%.2f, N=%.1f\nРаспределение на 4 подсистемы: %.2f тыс. руб. каждая",
			scale, zLinear, NLinear, zExp, NExp, Zs/4)
		resultLabel.SetText(res)

		p := plot.New()
		p.Title.Text = "Зависимость P(z)"
		p.X.Label.Text = "Затраты Z (тыс. руб.)"
		p.Y.Label.Text = "Срок службы P(z) × Масштаб"

		pointsLinear := make(plotter.XYs, 100)
		pointsExp := make(plotter.XYs, 100)
		for i := 0; i < 100; i++ {
			z := float64(i) * 0.5
			pointsLinear[i].X = z
			pointsLinear[i].Y = (k * z) * scale
			pointsExp[i].X = z
			pointsExp[i].Y = (Tbaz * (1 - math.Exp(-k*z))) * scale
		}

		lineLinear, _ := plotter.NewLine(pointsLinear)
		lineLinear.Color = color.RGBA{0, 0, 255, 255}
		lineExp, _ := plotter.NewLine(pointsExp)
		lineExp.Color = color.RGBA{255, 0, 0, 255}
		p.Add(lineLinear, lineExp)
		p.Legend.Add("Линейная", lineLinear)
		p.Legend.Add("Экспоненциальная", lineExp)

		Tline := plotter.NewFunction(func(x float64) float64 { return Tzadan * scale })
		Tline.Color = color.RGBA{0, 200, 0, 255}
		p.Add(Tline)
		p.Legend.Add("Tзадан", Tline)

		fileName := "plot.png"
		if err := p.Save(6*vg.Inch, 4*vg.Inch, fileName); err != nil {
			fmt.Println("Ошибка графика:", err)
			return
		}
		image.File = fileName
		image.Refresh()
	})
	
	reportButton := widget.NewButton("Отчёт", func() {
		saveReport()
	})

	refreshButton := widget.NewButton("Обновить", func() {
		image.File = ""
		image.Refresh()
		resultLabel.SetText("Здесь появятся результаты")
	})

	leftPanel := container.NewVBox(
		widget.NewLabel("Масштаб:"), scaleEntry,
		widget.NewLabel("Базовый срок службы:"), TbazEntry,
		widget.NewLabel("Нахождение затрат:"), TzadanEntry,
		widget.NewLabel("Распределение затрат:"), ZsEntry,
		calcButton,
		reportButton,
		refreshButton,
		resultLabel,
	)

	content := container.NewHSplit(leftPanel, image)
	content.SetOffset(0.3)

	myWindow.SetContent(content)
	myWindow.Resize(fyne.NewSize(1000, 600))
	myWindow.ShowAndRun()

	os.Remove("plot.png")
}