import pdfplumber  
import pandas as pd  
import os  
import glob  

# Файлы
excel_filename = "result.xlsx"
search_text = "Всего по акту"

# Функция для поиска всех PDF-файлов с "КС2" в имени
def find_pdf_files():
    file_pattern = os.path.join(os.getcwd(), '**', 'Печатная форма*КС2*.pdf')  # Шаблон поиска
    all_matches = glob.glob(file_pattern, recursive=True)
    return [f for f in all_matches if os.path.isfile(f)]  # Оставляем только файлы

# Находим все подходящие файлы
pdf_files = find_pdf_files()

if pdf_files:
    results = []  # Список для хранения результатов

    for pdf_filename in pdf_files:
        extracted_sum = None  # Сумма по текущему файлу

        try:
            with pdfplumber.open(pdf_filename) as pdf:
                # Перебираем страницы с конца
                for page in reversed(pdf.pages):  
                    tables = page.extract_tables()  # Извлекаем таблицы

                    for table in tables:
                        for row in table:
                            if row and row[0] and search_text in row[0]:  # Проверка "Всего по акту"
                                if len(row) > 8 and row[8]:  # Проверка, есть ли сумма
                                    raw_value = row[8].replace("\u00A0", "").replace(" ", "")  # Убираем пробелы и неразрывные пробелы
                                    try:
                                        extracted_sum = round(float(raw_value), 2)
                                    except ValueError:
                                        extracted_sum = raw_value   # Если вдруг не число, сохраняем как есть
                                    break  
                        if extracted_sum:
                            break
                    if extracted_sum:
                        break
        except Exception as e:
            extracted_sum = f"Ошибка: {str(e)}"

        # Добавляем результат
        results.append({
            "Файл": os.path.basename(pdf_filename),
            "Сумма": extracted_sum if extracted_sum else "Не найдено"
        })

    # Сохраняем в Excel
    df = pd.DataFrame(results)
    df.to_excel(excel_filename, index=False, engine="openpyxl")

    print(f"Обработано файлов: {len(results)}. Результаты сохранены в {excel_filename}.")
else:
    print("Не найдено ни одного PDF-файла с 'KC2' в имени.")