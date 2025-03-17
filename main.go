package main

import (
	"fmt"

	"github.com/carmel/gooxml/chart"
	"github.com/carmel/gooxml/measurement"
	"github.com/carmel/gooxml/presentation"
)

func main() {
	// Данные
	datas := map[string][]int{
		"Negative": {1, 2, 3, 4, 5},
		"Neutral":  {4, 6, 8, 5, 3},
		"Positive": {4, 7, 3, 5, 8},
		"max":      {9, 15, 14, 14, 16},
	}

	// Создаем новую презентацию
	ppt := presentation.New()
	slide := ppt.AddSlide()

	// Создаем график
	ch := chart.NewCombinationChart()
	ch.Properties().SetWidth(12 * measurement.Centimeter)
	ch.Properties().SetHeight(8 * measurement.Centimeter)

	// Добавляем бар для Negative
	bar := chart.NewBarChart()
	bar.Properties().SetBarDirection(chart.BarDirectionCol)

	negSeries := bar.AddSeries()
	negSeries.SetLabel("Negative")
	for _, v := range datas["Negative"] {
		negSeries.AddValue(fmt.Sprintf("%d", v), float64(v))
	}

	// Добавляем линию для max
	line := chart.NewLineChart()
	line.Properties().LineProperties().SetWidth(2 * measurement.Point)

	maxSeries := line.AddSeries()
	maxSeries.SetLabel("Max")
	for _, v := range datas["max"] {
		maxSeries.AddValue(fmt.Sprintf("%d", v), float64(v))
	}

	// Добавляем оба типа графиков в комбинированный график
	ch.AddChart(bar)
	ch.AddChart(line)

	// Настраиваем оси
	catAx := chart.NewCategoryAxis()
	valAx := chart.NewValueAxis()
	ch.AddAxis(catAx)
	ch.AddAxis(valAx)
	bar.SetCategoryAxis(catAx)
	bar.SetValueAxis(valAx)
	line.SetCategoryAxis(catAx)
	line.SetValueAxis(valAx)

	// Добавляем заголовок
	ch.Properties().Title().SetText("Negative vs Max Values")

	// Добавляем график на слайд
	slide.AddChart(ch, presentation.NewChartSpace().
		SetPosition(1*measurement.Centimeter, 1*measurement.Centimeter))

	// Сохраняем файл
	err := ppt.SaveToFile("chart.pptx")
	if err != nil {
		fmt.Printf("Ошибка при сохранении файла: %v\n", err)
		return
	}
	fmt.Println("График успешно создан и сохранен в chart.pptx")
}
