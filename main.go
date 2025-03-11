package main

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	// Получаем текущую рабочую директорию (где лежит exe файл)
	exeDir, err := os.Getwd()
	if err != nil {
		log.Fatalf("Ошибка при получении рабочей директории: %v", err)
	}

	// Ищем файл .xlsx только в текущем каталоге
	var excelFilePath string
	files, err := os.ReadDir(exeDir)
	if err != nil {
		log.Fatalf("Ошибка при чтении каталога: %v", err)
	}

	for _, file := range files {
		// Ищем файл с расширением .xlsx
		if !file.IsDir() && strings.HasSuffix(file.Name(), ".xlsx") {
			excelFilePath = filepath.Join(exeDir, file.Name())
			break
		}
	}

	if excelFilePath == "" {
		log.Fatalf("Excel файл не найден в каталоге: %v", exeDir)
	}

	// Открываем найденный Excel файл
	xlsx, err := excelize.OpenFile(excelFilePath)
	if err != nil {
		log.Fatalf("Ошибка при открытии файла: %v", err)
	}

	// Получаем имя первого листа
	sheetName := xlsx.GetSheetName(0) // Получаем имя первого листа, индексация начинается с 1
	if sheetName == "" {
		log.Fatalf("Ошибка: не удалось получить имя первого листа")
	}

	// Читаем все строки с листа
	rows, err := xlsx.GetRows(sheetName)
	if err != nil {
		log.Fatalf("Ошибка при чтении строк с листа '%s': %v", sheetName, err)
	}

	if len(rows) == 0 {
		log.Fatalf("Лист '%s' не содержит данных.", sheetName)
	}

	// Создаем каталог для обработанных файлов
	processedDir := filepath.Join(exeDir, "Обработанные")
	err = os.MkdirAll(processedDir, 0755)
	if err != nil {
		log.Fatalf("Ошибка при создании каталога Обработанные: %v", err)
	}

	// Карта для хранения писателей для каждого способа оплаты
	paymentFiles := make(map[string]*csv.Writer)

	// Читаем заголовки (первая строка)
	header := rows[0]

	// Создаем карту для хранения индексов столбцов по их названиям
	headerIndex := make(map[string]int)

	// Находим индексы столбцов для нужных заголовков
	for i, colName := range header {
		headerIndex[colName] = i
	}

	// Пропускаем заголовок (первая строка)
	for i, row := range rows {
		if i == 0 {
			continue
		}

		// Пропускаем строки с неполными данными
		if len(row) < 8 {
			continue
		}

		if row[headerIndex["Канал оплаты"]] != "Удаленная оплата" {
			continue
		}

		// Определяем новый способ оплаты
		paymentMethod := row[headerIndex["Способ оплаты"]]

		// Модификация способа оплаты для группировки
		if paymentMethod == "Kaspi Kredit" || paymentMethod == "Kaspi Red" || paymentMethod == "Кредит на Покупки" {
			paymentMethod = "Kaspi Red" // Объединяем Kaspi Kredit, Кредит на Покупки и Kaspi Red в Kaspi Red
		} else if paymentMethod == "Счет в Kaspi Pay" {
			paymentMethod = "Kaspi Gold"
		}

		// Извлекаем данные из строки
		re := regexp.MustCompile("[0-9]+")
		numbers := re.FindString(row[headerIndex["Детали покупки"]])

		// Создаем каталог для способа оплаты, если его еще нет
		paymentDir := filepath.Join(exeDir, paymentMethod)
		if _, err := os.Stat(paymentDir); os.IsNotExist(err) {
			err = os.MkdirAll(paymentDir, 0755)
			if err != nil {
				log.Fatalf("Ошибка при создании каталога для %s: %v", paymentMethod, err)
			}
		}

		// Преобразуем строку в число
		floatAmount, er := strconv.ParseFloat(strings.ReplaceAll(strings.ReplaceAll(row[headerIndex["Сумма операции"]], " ", ""), ",", ""), 64)
		if er != nil {
			fmt.Println("Ошибка преобразования:", er)
			return
		}

		// Форматируем результат без десятичных знаков
		amount := fmt.Sprintf("%.0f", floatAmount)

		// Форматируем строку для CSV
		data := []string{strings.Join([]string{numbers, row[headerIndex["Дата операции"]], amount, row[headerIndex["Номер операции"]], row[headerIndex["Способ оплаты"]]}, ",")}
		// Если для данного способа оплаты еще нет файла, создаем его
		if _, exists := paymentFiles[paymentMethod]; !exists {
			// Форматируем текущую дату и время
			currentTime := time.Now()
			timeFormatted := currentTime.Format("2006.01.02_15-04-05")
			fileName := fmt.Sprintf("%s_%s.csv", paymentMethod, timeFormatted)

			// Открываем новый CSV файл для этого способа оплаты
			filePath := filepath.Join(paymentDir, fileName)
			file, err := os.Create(filePath)
			if err != nil {
				log.Fatalf("Ошибка при создании CSV файла для %s: %v", paymentMethod, err)
			}
			defer file.Close()

			// Создаем новый писатель для этого файла
			paymentFiles[paymentMethod] = csv.NewWriter(file)

			// Добавляем BOM для правильной интерпретации UTF-8 в Excel
			_, err = file.Write([]byte{0xEF, 0xBB, 0xBF}) // Это байты для BOM
			if err != nil {
				log.Fatalf("Ошибка при записи BOM в файл: %v", err)
			}
		}

		// Записываем в соответствующий CSV файл
		err := paymentFiles[paymentMethod].Write(data)
		if err != nil {
			log.Printf("Ошибка записи в CSV файл для %s: %v", paymentMethod, err)
		}
	}

	// Закрываем все файлы
	for _, writer := range paymentFiles {
		writer.Flush()
	}

	// Перемещаем исходный файл в каталог "Обработанные"
	destPath := filepath.Join(processedDir, filepath.Base(excelFilePath))
	err = os.Rename(excelFilePath, destPath)
	if err != nil {
		log.Fatalf("Ошибка перемещения файла: %v", err)
	}

	fmt.Println("Парсинг завершен успешно!")
}
