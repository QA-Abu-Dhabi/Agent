import pdfplumber  
import pandas as pd  

# Файлы
pdf_filename = "document.pdf"
excel_filename = "result.xlsx"

# Искомая фраза
search_text = "Всего по акту"

# Открываем PDF
with pdfplumber.open(pdf_filename) as pdf:
    extracted_sum = None  # Переменная для суммы
    
    # Перебираем страницы
    for page in pdf.pages:
        tables = page.extract_tables()  # Извлекаем таблицы
        
        for table in tables:
            for row in table:
                if row and row[0] and search_text in row[0]:  # Проверяем, есть ли "Всего по акту" в первой ячейке
                    if len(row) > 8 and row[8]:  # Проверяем, есть ли сумма в 9-й колонке (индекс 8)
                        extracted_sum = row[8].replace(" ", "")  # Очищаем сумму от пробелов
                        break  
            if extracted_sum:
                break
        if extracted_sum:
            break

# Если сумма найдена, записываем в Excel
if extracted_sum:
    df = pd.DataFrame({"Сумма": [extracted_sum]})  # Создаем DataFrame
    df.to_excel(excel_filename, index=False, engine="openpyxl")  # Сохраняем
    print(f"Сумма '{extracted_sum}' сохранена в {excel_filename}")
else:
    print("Текст 'Всего по акту' не найден в таблицах PDF.")