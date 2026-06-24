package main

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"math"
	"net/http"
	"os/exec"
	"strconv"
	"strings"
	"time"

	"github.com/lukasjarosch/go-docx" // библиотека для работы с docx
	"github.com/xuri/excelize/v2"     // библиотека для работы с xlsx
)

// Структура, которая совпадает с ключами моего json от ИИ (для шаблона ВОРД акта(АОСР))
type StampInfo struct {
	ProjectCode string `json:"ProjectCode"` //шифр
	SheetNumber string `json:"SheetNumber"` //номер листа
	WorkType    string `json:"WorkType"`    //вид работ
	Axes        string `json:"Axes"`        //оси
	Floor       string `json:"Floor"`       //этаж
	Block       string `json:"Block"`       //блок
	ShemaNumber string `json:"ShemaNumber"` //номер схемы
}

// --- СТРУКТУРЫ ДЛЯ РАЗБОРА ОТВЕТОВ ЯНДЕКСА ---

// VisionResponse — структура для парсинга ответа от Yandex Vision API
type VisionResponse struct {
	Results []struct {
		Results []struct {
			TextDetection struct {
				Pages []struct {
					Blocks []struct {
						Lines []struct {
							Words []struct {
								Text string `json:"text"`
							} `json:"words"`
						} `json:"lines"`
					} `json:"blocks"`
				} `json:"pages"`
			} `json:"textDetection"`
		} `json:"results"`
	} `json:"results"`
}

// GptResponse — структура для парсинга ответа от Yandex GPT
type GptResponse struct {
	Result struct {
		Alternatives []struct {
			Message struct {
				Text string `json:"text"`
			} `json:"message"`
		} `json:"alternatives"`
	} `json:"result"`
}

// --- ОСНОВНАЯ ЛОГИКА ---
func main() {
	// Начало отсчёта времени выполнения
	start := time.Now()

	// НАСТРОЙКИ подключения 
	apiKey := " ваш apiKey" 
	folderID := "ваш folderID"

	pdfPath := "scan.pdf" //схема в пдф, из которой берем данные

	fmt.Println("=== СИСТЕМА СТРОЙ-ИИ: ЗАПУСК ===")

	// Проверка формата PDF: А0/А1 или стандартный?
	fmt.Println("=== Шаг 1: Анализ формата PDF... ===")
	isLarge, width, height := checkPDFSize(pdfPath)

	if isLarge {
		// Большой чертёж: обрабатываем зону штампа
		fmt.Printf(">> Обнаружен формат А1/А0 (%.0fx%.0f pts). Запуск обработки зоны штампа...\n", width, height)

		// Коэффициент для 300 DPI (300 / 72 = 4.166)
		const ratio = 4.166
		Wpx := width * ratio
		Hpx := height * ratio

		// Плитка 60% от размера листа
		tileW := Wpx * 0.6
		tileH := Hpx * 0.6

		// Координаты зоны штампа (Право-Низ)
		zones := [][]float64{
			{Wpx - tileW, Hpx - tileH, tileW, tileH},
		}

		var fullPageText []string

		for i, z := range zones {
			fmt.Printf(">> Обработка зоны штампа %d...\n", i+1)
			cmd := exec.Command("pdftoppm", "-jpeg", "-singlefile", "-r", "300",
				"-x", strconv.FormatFloat(z[0], 'f', 0, 64),
				"-y", strconv.FormatFloat(z[1], 'f', 0, 64),
				"-W", strconv.FormatFloat(z[2], 'f', 0, 64),
				"-H", strconv.FormatFloat(z[3], 'f', 0, 64),
				pdfPath)

			var outBuf bytes.Buffer
			cmd.Stdout = &outBuf
			if err := cmd.Run(); err == nil {
				text := getRawOCR(outBuf.Bytes(), apiKey, folderID)
				fullPageText = append(fullPageText, text)
			}
		}

		finalRawText := strings.Join(fullPageText, " ")

		fmt.Println("=== ШАГ 2: ИНТЕЛЛЕКТУАЛЬНЫЙ АНАЛИЗ ===")

		// В блоке, где формируется prompt для Yandex
		prompt := fmt.Sprintf(`Ты - эксперт ПТО, анализируешь чертежи.
ИНСТРУКЦИИ:
1. ИЗ ШТАМПА: Найди и выдели следующие поля:
 1. Шифр (ProjectCode): это уникальный индекс проекта. Обычно собержит цифры, дефисы и скоращения (например, ИМ-20-7052-АС, 12/23-КЖ, 2024-ИД-15). Ищи его в нижней правой части.
 2. Номер листа (SheetNumber): это номер листа, на котором находится чертеж. Обычно это число, чтоит рядом с "Лист" или в графе "Лист". Это число расположено в рамке справа от слова "ИД" или "РД"
 3. Вид работ (WorkType): чаще всего в виде работ не пишется название компании, то есть длинного слова с заглавными буквами чаще быть не должно. Вот пример вида работ: "Монтаж диффузоров, адаптеров с жёсткими и гибкими горизонтальными участками воздуховодов"
 4. Оси (Axes)
 5. Этаж (Floor)
 6. Блок/Секция (Block)
 7. Номер исполнительной схемы (ShemaNumber)
   

ФОРМАТ ОТВЕТА:
- Только чистый JSON.
- Пример структуры для штампа:
  "stamp": {
    "cipher": "АБВГ.123456.001",
    "sheet": "1",
    "title": "Схема расположения оборудования",
    "axes": "1-5/А-Г"
  }

ТЕКСТ ДЛЯ АНАЛИЗА:
%s`, finalRawText)

		payload := map[string]interface{}{
			"modelUri":          "gpt://" + folderID + "/yandexgpt/latest",
			"completionOptions": map[string]interface{}{"stream": false, "temperature": 0.1, "maxTokens": 8000},
			"messages": []map[string]interface{}{
				{"role": "system", "text": "Ты эксперт ПТО. Отвечай только чистым JSON."},
				{"role": "user", "text": prompt},
			},
		}

		// Отправляем запрос к Yandex GPT API на генерацию ответа модели
		gptJSON, _ := json.Marshal(payload)
		req, _ := http.NewRequest("POST", "https://llm.api.cloud.yandex.net/foundationModels/v1/completion", bytes.NewBuffer(gptJSON))
		req.Header.Set("Content-Type", "application/json")
		req.Header.Set("Authorization", "Api-Key "+apiKey)

		//Проверка на ошибки
		client := &http.Client{}
		resp, err := client.Do(req)
		if err != nil {
			fmt.Printf("Ошибка запроса: %v\n", err)
			return
		}
		defer resp.Body.Close()

		body, err := io.ReadAll(resp.Body)
		if err != nil {
			fmt.Printf("Ошибка чтения тела: %v\n", err)
			return
		}

		var gRes GptResponse
		err = json.Unmarshal(body, &gRes)
		if err != nil {
			fmt.Printf("Ошибка Парсинга: %v\nСырой JSON был: %s\n", err, string(body))
			return
		}

		if len(gRes.Result.Alternatives) == 0 {
			fmt.Println("Ошибка: ИИ прислал пустой список альтернатив")
			return
		}

		fmt.Println("\nФИНАЛЬНЫЙ ОТЧЕТ (JSON):")
		if len(gRes.Result.Alternatives) > 0 {
			gptResponseText := gRes.Result.Alternatives[0].Message.Text
			fmt.Println(gptResponseText)

			//данные для БОЛЬШОГО ФОРМАТА(А0, А1, А2) схемы
			cleanedJSON := cleanJSONResponse(gptResponseText)
			templateFile := "5full_act_white.docx" //шаблон для акта
			excelFile := "const.xlsx"  //данные константы
			outputFile := "5test_out_full_act_big.docx" //наименования готового акта
			fillWordTemplate(cleanedJSON, excelFile, templateFile, outputFile)
		}

	} else {
		// Стандартный формат
		fmt.Println(">> Стандартный формат. Обработка целиком.")
		cmd := exec.Command("pdftoppm", "-jpeg", "-singlefile", "-r", "300", pdfPath)

		fmt.Println(">> Шаг 1: Растеризация чертежа...")
		var outBuffer bytes.Buffer
		cmd.Stdout = &outBuffer
		if err := cmd.Run(); err != nil {
			fmt.Printf("Ошибка pdftoppm: %v\n", err)
			return
		}

		fmt.Println(">> Шаг 2: Распознавание текста (OCR)...")
		encodedFile := base64.StdEncoding.EncodeToString(outBuffer.Bytes())

		visionPayload := map[string]interface{}{
			"folderId": folderID,
			"analyzeSpecs": []map[string]interface{}{
				{
					"content": encodedFile,
					"features": []map[string]interface{}{
						{"type": "TEXT_DETECTION", "textDetectionConfig": map[string]interface{}{"languageCodes": []string{"ru", "en"}}},
					},
				},
			},
		}

		// Отправляем запрос к Yandex GPT API
		visionJSON, _ := json.Marshal(visionPayload)
		reqV, _ := http.NewRequest("POST", "https://vision.api.cloud.yandex.net/vision/v1/batchAnalyze", bytes.NewBuffer(visionJSON))
		reqV.Header.Set("Content-Type", "application/json")
		reqV.Header.Set("Authorization", "Api-Key "+apiKey)

		//Проверка на ошибки
		client := &http.Client{}
		respV, err := client.Do(reqV)
		if err != nil {
			fmt.Printf("Ошибка Vision: %v\n", err)
			return
		}
		defer respV.Body.Close()

		bodyV, _ := io.ReadAll(respV.Body)
		var vRes VisionResponse
		json.Unmarshal(bodyV, &vRes)

		var words []string
		for _, resO := range vRes.Results {
			for _, resI := range resO.Results {
				for _, page := range resI.TextDetection.Pages {
					for _, block := range page.Blocks {
						for _, line := range block.Lines {
							for _, word := range line.Words {
								words = append(words, word.Text)
							}
						}
					}
				}
			}
		}
		fullText := strings.Join(words, " ")

		if len(fullText) < 10 {
			fmt.Println("Текст не найден.")
			return
		}

		fmt.Println(">> Шаг 3: Интеллектуальный анализ штампа...")

		// В блоке, где формируется prompt для Yandex
		prompt := fmt.Sprintf(`Ты - эксперт ПТО, анализируешь чертежи.
ИНСТРУКЦИИ:
1. ИЗ ШТАМПА: Найди и выдели следующие поля:
 1. Шифр (ProjectCode): это уникальный индекс проекта. Обычно собержит цифры, дефисы и скоращения (например, ИМ-20-7052-АС, 12/23-КЖ, 2024-ИД-15). Ищи его в нижней правой части.
 2. Номер листа (SheetNumber): это номер листа, на котором находится чертеж. Обычно это число, чтоит рядом с "Лист" или в графе "Лист". Это число расположено в рамке справа от слова "ИД" или "РД"
 3. Вид работ (WorkType): чаще всего в виде работ не пишется название компании, то есть длинного слова с заглавными буквами чаще быть не должно. Вот пример вида работ: "Монтаж диффузоров, адаптеров с жёсткими и гибкими горизонтальными участками воздуховодов"
 4. Оси (Axes)
 5. Этаж (Floor)
 6. Блок/Секция (Block)
 7. Номер исполнительной схемы (ShemaNumber)
   

ФОРМАТ ОТВЕТА:
- Только чистый JSON.
- Пример структуры для штампа:
  "stamp": {
    "cipher": "АБВГ.123456.001",
    "sheet": "1",
    "title": "Схема расположения оборудования",
    "axes": "1-5/А-Г"
  }

ТЕКСТ ДЛЯ АНАЛИЗА:
%s`, fullText)

		gptPayload := map[string]interface{}{
			"modelUri":          "gpt://" + folderID + "/yandexgpt/latest",
			"completionOptions": map[string]interface{}{"stream": false, "temperature": 0.1, "maxTokens": 1000},
			"messages": []map[string]interface{}{
				{"role": "system", "text": "Ты эксперт строительной документации. Отвечай только чистым JSON."},
				{"role": "user", "text": prompt},
			},
		}
		// Отправляем запрос к Yandex GPT API на генерацию ответа модели
		gptJSON, _ := json.Marshal(gptPayload)
		reqG, _ := http.NewRequest("POST", "https://llm.api.cloud.yandex.net/foundationModels/v1/completion", bytes.NewBuffer(gptJSON))
		reqG.Header.Set("Content-Type", "application/json")
		reqG.Header.Set("Authorization", "Api-Key "+apiKey)

		respG, err := client.Do(reqG)
		if err != nil {
			fmt.Printf("Ошибка: %v\n", err)
			return
		}
		defer respG.Body.Close()

		bodyG, _ := io.ReadAll(respG.Body)
		var gRes GptResponse
		json.Unmarshal(bodyG, &gRes)

		fmt.Println("\n=== ИТОГОВЫЙ JSON ДЛЯ АОСР ===")
		if len(gRes.Result.Alternatives) > 0 {
			gptResponseText := gRes.Result.Alternatives[0].Message.Text
			fmt.Println("Ответ GPT:", gptResponseText)

			//данные для МАЛЕНЬКОГО ФОРМАТА(А3, А4) схемы
			cleanedJSON := cleanJSONResponse(gptResponseText)
			templateFile := "test_full_act.docx" //шаблон акта
			excelFile := "const.xlsx"  //данные константы
			outputFile := "test_out_full_small.docx" //нименование готового акта

			fillWordTemplate(cleanedJSON, excelFile, templateFile, outputFile)
		}
	}

	fmt.Printf("Время выполнения: %.2f секунд\n", time.Since(start).Seconds()) //время работы программы
}

// --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

// checkPDFSize - определяет формат PDF
func checkPDFSize(path string) (bool, float64, float64) {
	out, err := exec.Command("pdfinfo", path).Output()
	if err != nil {
		return false, 0, 0
	}
	lines := strings.Split(string(out), "\n")
	for _, line := range lines {
		if strings.Contains(line, "Page size:") {
			parts := strings.Fields(line)
			if len(parts) >= 5 {
				w, _ := strconv.ParseFloat(parts[2], 64)
				h, _ := strconv.ParseFloat(parts[4], 64)
				return math.Max(w, h) > 1700, w, h
			}
		}
	}
	return false, 0, 0
}

// getRawOCR - отправляет изображение в Yandex Vision
func getRawOCR(imgBytes []byte, apiKey, folderID string) string {
	encoded := base64.StdEncoding.EncodeToString(imgBytes)
	payload := map[string]interface{}{
		"folderId": folderID,
		"analyzeSpecs": []map[string]interface{}{{
			"content": encoded,
			"features": []map[string]interface{}{{
				"type":                "TEXT_DETECTION",
				"textDetectionConfig": map[string]interface{}{"languageCodes": []string{"ru", "en"}},
			}},
		}},
	}
	jsonData, _ := json.Marshal(payload)
	req, _ := http.NewRequest("POST", "https://vision.api.cloud.yandex.net/vision/v1/batchAnalyze", bytes.NewBuffer(jsonData))
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("Authorization", "Api-Key "+apiKey)

	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		return ""
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)

	var vRes VisionResponse
	json.Unmarshal(body, &vRes)

	var fullTextBuilder strings.Builder
	for _, res0 := range vRes.Results {
		for _, res1 := range res0.Results {
			for _, page := range res1.TextDetection.Pages {
				for _, block := range page.Blocks {
					for _, line := range block.Lines {
						var lineWords []string
						for _, word := range line.Words {
							lineWords = append(lineWords, word.Text)
						}
						if len(lineWords) > 0 {
							fullTextBuilder.WriteString(strings.Join(lineWords, " "))
							fullTextBuilder.WriteString("\n")
						}
					}
				}
			}
		}
	}
	return fullTextBuilder.String()
}

// fillWordTemplate - заполняет Word шаблон данными из JSON и Excel
func fillWordTemplate(jsonStr string, excelPath string, templatePath string, outputPath string) {
	cleanedJSON := cleanJSONResponse(jsonStr)

	var rawData map[string]interface{}
	if err := json.Unmarshal([]byte(cleanedJSON), &rawData); err != nil {
		log.Printf("Ошибка парсинга JSON: %v", err)
		return
	}

	stampData, ok := rawData["stamp"].(map[string]interface{})
	if !ok {
		log.Printf("Ошибка: нет поля 'stamp' в JSON")
		return
	}

	// Загружаем данные из Excel
	excelData := make(map[string]string)
	if excelPath != "" {
		excelData, _ = readConstantsFromExcel(excelPath, "")
	}

	// Создаем карту замены (ключи БЕЗ фигурных скобок!)
	replaceMap := make(docx.PlaceholderMap)
	replaceMap["ProjectCode"] = getStringValue(stampData["cipher"])
	replaceMap["SheetNumber"] = getStringValue(stampData["sheet"])
	replaceMap["WorkType"] = getStringValue(stampData["title"])
	replaceMap["Axes"] = getStringValue(stampData["axes"])
	replaceMap["Floor"] = getStringValue(stampData["floor"])
	replaceMap["Block"] = getStringValue(stampData["block"])
	replaceMap["ShemaNumber"] = getStringValue(stampData["shemaNumber"])

	for k, v := range excelData {
		replaceMap[k] = v
	}

	// Заполняем пустые значения
	for key, value := range replaceMap {
		if value == "" {
			replaceMap[key] = "___________"
		}
	}

	fmt.Println("\n📋 Данные для замены:")
	for k, v := range replaceMap {
		// Приводим к строке безопасно
		strValue := fmt.Sprintf("%v", v)
		if len(strValue) > 60 {
			fmt.Printf("  ├─ %s: %s...\n", k, strValue[:57])
		} else {
			fmt.Printf("  ├─ %s: %s\n", k, strValue)
		}
	}

	// Открываем и обрабатываем шаблон
	doc, err := docx.Open(templatePath)
	if err != nil {
		log.Fatalf("Ошибка открытия шаблона: %v", err)
	}
	defer doc.Close()

	fmt.Println("\n🔄 Выполняется замена меток...")
	err = doc.ReplaceAll(replaceMap)
	if err != nil {
		log.Fatalf("Ошибка замены: %v", err)
	}
	fmt.Println("  ✅ Замена выполнена успешно")

	// Сохраняем результат
	err = doc.WriteToFile(outputPath)
	if err != nil {
		log.Fatalf("Ошибка сохранения: %v", err)
	}

	fmt.Printf("\n✅ WORD-файл успешно создан: %s\n", outputPath)
}

// getStringValue - безопасно преобразует interface{} в string
func getStringValue(v interface{}) string {
	if v == nil {
		return ""
	}
	switch val := v.(type) {
	case string:
		return val
	case float64:
		return fmt.Sprintf("%g", val)
	case int:
		return fmt.Sprintf("%d", val)
	default:
		return fmt.Sprintf("%v", val)
	}
}

// cleanJSONResponse - очищает JSON от лишних символов
func cleanJSONResponse(raw string) string {
	raw = strings.TrimSpace(raw)

	if strings.HasPrefix(raw, "```json") {
		raw = strings.TrimPrefix(raw, "```json")
		raw = strings.TrimSuffix(raw, "```")
	} else if strings.HasPrefix(raw, "```") {
		raw = strings.TrimPrefix(raw, "```")
		raw = strings.TrimSuffix(raw, "```")
	}

	raw = strings.TrimSpace(raw)

	start := strings.Index(raw, "{")
	end := strings.LastIndex(raw, "}")

	if start != -1 && end != -1 && end > start {
		return raw[start : end+1]
	}

	return raw
}

// readConstantsFromExcel - читает данные из Excel
func readConstantsFromExcel(excelPath, sheetName string) (map[string]string, error) {
	f, err := excelize.OpenFile(excelPath)
	if err != nil {
		return nil, fmt.Errorf("ошибка открытия Excel: %w", err)
	}
	defer f.Close()

	if sheetName == "" {
		sheetName = f.GetSheetName(0)
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("ошибка чтения строк: %w", err)
	}

	if len(rows) < 2 {
		return nil, fmt.Errorf("Excel файл пуст или нет данных")
	}

	dataMap := make(map[string]string)

	for i, row := range rows {
		if i == 0 {
			continue
		}
		if len(row) >= 2 {
			key := strings.TrimSpace(row[0])
			value := strings.TrimSpace(row[1])
			if key != "" {
				dataMap[key] = value
				fmt.Printf("  ├─ Загружено: %s = %s\n", key, value)
			}
		}
	}

	return dataMap, nil
}
