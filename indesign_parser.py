#!/usr/bin/env python3
import os
import sys
import glob
import pandas as pd

def process_excel_to_indesign(file_path):
    print(f"Обработка файла: {os.path.basename(file_path)}")
    
    try:
        # dtype=str жестко заставляет pandas читать всё как текст, не превращая "1" в "1.0"
        df = pd.read_excel(file_path, header=None, engine='calamine', dtype=str)
    except Exception as e:
        print(f"Ошибка при чтении {file_path}: {e}")
        return

    df = df.fillna("")
    
    orders_data = {}
    quotes_data = {}
    current_order_id = None

    # Построчное чтение
    for row in df.values.tolist():
        # Зачищаем данные: убираем пробелы по краям и срезаем ".0", если они проскочили
        row_vals = []
        for cell in row:
            val = str(cell).strip()
            if val.endswith('.0'):
                val = val[:-2]
            row_vals.append(val)
            
        if not any(row_vals):
            continue
        
        # --- БЛОК 1: Сбор заказов ---
        # Проверяем, что первый столбец — это число (номер п/п), а во втором есть "№:"
        if row_vals[0].isdigit() and "№:" in row_vals[1]:
            # Достаем чистый ID заказа
            order_id = row_vals[1].split(',')[0].replace("№:", "").strip()
            date_str = row_vals[2]
            raw_name = row_vals[3]
            
            # ДОБАВЛЕНО: Извлекаем статус оплаты заказа напрямую из строки
            is_paid = "Оплачен" in row_vals[1]
            
            photos_raw = row_vals[8] if len(row_vals) > 8 else "" 
            
            # Всеядный парсинг даты через pandas
            try:
                parsed_date = pd.to_datetime(date_str, dayfirst=True)
            except Exception:
                parsed_date = pd.Timestamp.min
            
            # Сохраняем, только если это новый заказ или дата свежее
            if order_id not in orders_data or parsed_date > orders_data.get(order_id, {}).get('parsed_date', pd.Timestamp.min):
                orders_data[order_id] = {
                    "name": raw_name,
                    "photos": photos_raw,
                    "parsed_date": parsed_date,
                    "is_paid": is_paid  # Сохраняем статус оплаты для последующей фильтрации
                }
                
        # --- БЛОК 2: Сбор цитат ---
        # Определяем строку начала блока цитат по длине ID заказа (> 4 символов).
        if row_vals[0].isdigit() and len(row_vals[0]) > 4:
            current_order_id = row_vals[0]
            
        if "Большое текстовое поле" in row_vals:
            idx = row_vals.index("Большое текстовое поле")
            if len(row_vals) > idx + 1:
                # Убираем кавычки и заменяем внутренние переносы строк на пробелы, 
                # чтобы текст цитаты гарантированно оставался на одной строчке
                quote = row_vals[idx + 1].strip('"').replace('\n', ' ').replace('\r', ' ')
                if current_order_id and quote:
                    quotes_data[current_order_id] = quote

    # --- Подготовка итоговых данных ---
    header = ["Фамилия", "Имя", " @Фото1", " @Фото2", " @Фото3", " @Фото4", " @Фото5", " @Фото6", " @Фото7", " @Фото8", "Цитата"]
    users_final = {}
    
    for order_id, data in orders_data.items():
        name_parts = data['name'].split()
        if not name_parts:
            continue
            
        last_name = name_parts[0].capitalize()
        first_name = name_parts[1].capitalize() if len(name_parts) > 1 else ""
        
        user_key = f"{last_name}_{first_name}"
        
        photos = data['photos'].split()
        formatted_photos = [f"{p}.jpg" for p in photos if p.startswith("ART_") or p.startswith("ZUZ_")]
        
        # Добиваем пустые ячейки до 8 штук
        while len(formatted_photos) < 8:
            formatted_photos.append("")
        formatted_photos = formatted_photos[:8]
        
        quote = quotes_data.get(order_id, "")
        
        # ИЗМЕНЕНО: Новая каскадная логика фильтрации дубликатов по правилам оплаты и дат
        overwrite = False
        if user_key not in users_final:
            overwrite = True
        else:
            existing = users_final[user_key]
            # 1. Если новый заказ оплачен, а старый сохраненный нет — берем оплаченный
            if data['is_paid'] and not existing['is_paid']:
                overwrite = True
            # 2. Если новый не оплачен, а старый оплачен — игнорируем новый
            elif not data['is_paid'] and existing['is_paid']:
                overwrite = False
            # 3. Если оба оплачены или оба НЕ оплачены — берем тот, что позже по дате
            else:
                if data['parsed_date'] > existing['parsed_date']:
                    overwrite = True
                    
        if overwrite:
            users_final[user_key] = {
                "last_name": last_name,
                "first_name": first_name,
                "photos": formatted_photos,
                "quote": quote,
                "parsed_date": data['parsed_date'],
                "is_paid": data['is_paid']  # Сохраняем статус для последующих сравнений
            }

    # --- Экспорт в файл ---
    output_filename = os.path.splitext(file_path)[0] + ".txt"
    
    with open(output_filename, 'w', encoding='utf-16') as f:
        f.write("\t".join(header) + "\n")
        
        for key in sorted(users_final.keys()):
            u = users_final[key]
            line_elements = [u['last_name'], u['first_name']] + u['photos'] + [u['quote']]
            f.write("\t".join(line_elements) + "\n")
            
    print(f"Готово! Сгенерировано персон: {len(users_final)}")
    print(f"Файл сохранен: {os.path.basename(output_filename)}\n")

def main():
    target_dir = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
    
    if not os.path.isdir(target_dir):
        print(f"Ошибка: Директория '{target_dir}' не найдена.")
        sys.exit(1)
        
    search_pattern = os.path.join(target_dir, '*.xlsx')
    excel_files = glob.glob(search_pattern)
    
    excel_files = [f for f in excel_files if not os.path.basename(f).startswith('~$')]
    
    if not excel_files:
        print(f"Файлы .xlsx не найдены в папке {target_dir}")
        return
        
    for file_path in excel_files:
        process_excel_to_indesign(file_path)

    input("Нажмите Enter, чтобы закрыть окно...")

if __name__ == "__main__":
    main()