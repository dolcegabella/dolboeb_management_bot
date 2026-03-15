import telebot
import os
import cv2
import pandas as pd
from datetime import datetime
import easyocr
import shutil
import time
import re
import threading

# Токен бота
TOKEN = '8697524461:AAFs8il54OBoGjs8VnrvoQGkgplvxuYUDZ8'
bot = telebot.TeleBot(TOKEN)

# Создаем папки
for folder in ['temp', 'backups']:
    if not os.path.exists(folder):
        os.makedirs(folder)

# Инициализируем EasyOCR
print("Загрузка моделей OCR...")
reader = easyocr.Reader(['ru', 'en'], gpu=False)
print("Готов к работе!")

EXCEL_FILE = 'ебанаты (кроме кисули).xlsx'

# Слова для пропуска строк (приводим к нижнему регистру для надежности)
SKIP_WORDS = [word.lower() for word in ['участник', 'поиск', 'организатор', 'конференции', 'участники', 'Губернаторов', "Тимофеева", "Горбатенко", "Никифорова", "Тееленко", "Крылов"]]

# Словарь для хранения временных данных медиа-групп
media_groups = {}

def make_backup():
    """Создает резервную копию Excel файла"""
    if os.path.exists(EXCEL_FILE):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"backups/ебанаты_backup_{timestamp}.xlsx"
        shutil.copy2(EXCEL_FILE, backup_name)
        return backup_name
    return None

def clean_text(text):
    """
    Очищает текст, сохраняя пробелы
    """
    # Заменяем множественные пробелы на один
    text = re.sub(r'\s+', ' ', text)
    # Удаляем пробелы в начале и конце
    text = text.strip()
    return text

def remove_digits(text):
    """
    Удаляет все цифры из текста
    """
    return re.sub(r'\d+', '', text)

def filter_short_words(text):
    """
    Фильтрует слова, оставляя только те, что длиннее 2 символов (минимум 3 символа)
    """
    if not text:
        return text
    
    # Разбиваем на слова
    words = text.split()
    
    # Фильтруем слова (оставляем только те, где > 2 символов, т.е. минимум 3)
    filtered_words = [word for word in words if len(word) > 2]
    
    # Соединяем обратно через пробелы
    return ' '.join(filtered_words)

def capitalize_words(text):
    """
    Делает первую букву каждого слова заглавной, остальные строчные
    """
    if not text:
        return text
    
    # Разбиваем на слова по пробелам
    words = text.split()
    
    # Каждое слово с заглавной буквы
    capitalized_words = []
    for word in words:
        if len(word) > 0:
            # Первая буква заглавная, остальные строчные
            capitalized = word[0].upper() + word[1:].lower()
            capitalized_words.append(capitalized)
    
    # Соединяем обратно через пробелы
    return ' '.join(capitalized_words)

def should_skip_line(line_text):
    """
    Проверяет, нужно ли пропустить строку
    """
    line_lower = line_text.lower()
    
    # ВАЖНО: строки с этими фамилиями НИКОГДА не пропускаем
    important_names = ['ситдикова', 'безруков', 'мадина']
    for name in important_names:
        if name in line_lower:
            return False
    
    # Для остальных строк проверяем стоп-слова
    for word in SKIP_WORDS:
        if word in line_lower:
            print(f"Пропущена строка (содержит '{word}'): {line_text}")
            return True
    return False

def apply_special_rules(text):
    """
    Применяет специальные правила для определенных имен
    """
    if not text:
        return text
    
    original_text = text
    print(f"🔍 ПРИМЕНЯЮ ПРАВИЛА К: '{text}'")
    
    # ПРАВИЛО 1: Ситдикова - просто удаляем (попандо)
    if 'Ситдикова' in text or 'ситдикова' in text or 'СИТДИКОВА' in text:
        # Убираем (попандо) если есть
        text = re.sub(r'\(?попандо\)?', '', text, flags=re.IGNORECASE)
        text = re.sub(r'\(?popando\)?', '', text, flags=re.IGNORECASE)
        text = re.sub(r'\s+', ' ', text).strip()
        print(f"✅ ПРАВИЛО СИТДИКОВА (удален попандо): '{original_text}' -> '{text}'")
    
    # ПРАВИЛО 2: Мадина -> Мадина Задворочнова
    if 'Мадина' in text or 'мадина' in text or 'МАДИНА' in text:
        if 'Задворочнова' not in text and 'задворочнова' not in text:
            text = text + ' Задворочнова'
            print(f"✅ ПРАВИЛО МАДИНА: '{original_text}' -> '{text}'")
    
    # ПРАВИЛО 3: Безруков -> убираем Романович
    if 'Безруков' in text or 'безруков' in text or 'БЕЗРУКОВ' in text:
        text = re.sub(r'романович', '', text, flags=re.IGNORECASE)
        text = re.sub(r'\s+', ' ', text).strip()
        print(f"✅ ПРАВИЛО БЕЗРУКОВ: '{original_text}' -> '{text}'")
    
    return text

def extract_names_from_image(image_path):
    """
    Извлекает текст с фото и разбивает на строки
    """
    # Читаем и увеличиваем изображение
    img = cv2.imread(image_path)
    scale_percent = 150
    width = int(img.shape[1] * scale_percent / 100)
    height = int(img.shape[0] * scale_percent / 100)
    img = cv2.resize(img, (width, height), interpolation=cv2.INTER_CUBIC)
    
    # Конвертируем в серый
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Временно сохраняем
    temp_path = image_path.replace('.jpg', '_temp.jpg')
    cv2.imwrite(temp_path, gray)
    
    # Распознаем текст
    results = reader.readtext(temp_path, paragraph=False)
    
    # Удаляем временный файл
    try:
        os.remove(temp_path)
    except:
        pass
    
    if not results:
        return []
    
    # Выводим все распознанные тексты для отладки
    print("🔍 OCR РАСПОЗНАЛ:")
    for (bbox, text, confidence) in results:
        print(f"   Текст: '{text}', уверенность: {confidence:.2f}")
    
    # Группируем по строкам на основе координат
    words_with_pos = []
    for (bbox, text, confidence) in results:
        if confidence > 0.2:
            # Получаем координаты центра
            y_center = (bbox[0][1] + bbox[2][1]) / 2
            x_center = (bbox[0][0] + bbox[1][0]) / 2
            
            # Очищаем текст
            cleaned_text = clean_text(text)
            
            if cleaned_text and len(cleaned_text) >= 1:
                words_with_pos.append({
                    'text': cleaned_text,
                    'y': y_center,
                    'x': x_center
                })
    
    if not words_with_pos:
        return []
    
    # Сортируем по вертикали
    words_with_pos.sort(key=lambda x: x['y'])
    
    # Группируем по строкам
    lines = []
    current_line = []
    current_y = words_with_pos[0]['y']
    y_threshold = 25  # Порог для определения одной строки
    
    for word in words_with_pos:
        if abs(word['y'] - current_y) < y_threshold:
            current_line.append(word)
        else:
            # Сортируем слова в строке по горизонтали
            current_line.sort(key=lambda w: w['x'])
            
            # Собираем текст строки
            line_words = [w['text'] for w in current_line]
            line_text = ' '.join(line_words)
            
            # Очищаем от лишних пробелов
            line_text = clean_text(line_text)
            
            lines.append(line_text)
            current_line = [word]
            current_y = word['y']
    
    # Добавляем последнюю строку
    if current_line:
        current_line.sort(key=lambda w: w['x'])
        line_words = [w['text'] for w in current_line]
        line_text = ' '.join(line_words)
        line_text = clean_text(line_text)
        lines.append(line_text)
    
    print("🔍 СФОРМИРОВАННЫЕ СТРОКИ:")
    for i, line in enumerate(lines):
        print(f"   Строка {i+1}: '{line}'")
    
    # Обрабатываем каждую строку
    names = []
    
    for line in lines:
        # Проверяем, нужно ли пропустить строку
        if should_skip_line(line):
            continue
        
        print(f"🔍 ОБРАБАТЫВАЮ СТРОКУ: '{line}'")
        
        # ШАГ 1: Удаляем цифры из строки
        line_without_digits = remove_digits(line)
        if line_without_digits != line:
            print(f"   После удаления цифр: '{line_without_digits}'")
        
        # ШАГ 2: Применяем специальные правила (Ситдикова, Мадина, Безруков)
        line_with_rules = apply_special_rules(line_without_digits)
        
        # ШАГ 3: Делаем первую букву каждого слова заглавной
        formatted_line = capitalize_words(line_with_rules)
        
        # ШАГ 4: Очищаем от лишних пробелов
        formatted_line = clean_text(formatted_line)
        
        # ШАГ 5: Фильтруем короткие слова (оставляем только слова >= 3 букв)
        filtered_line = filter_short_words(formatted_line)
        print(f"   После фильтрации коротких: '{filtered_line}'")
        
        # Добавляем только непустые строки
        if filtered_line and len(filtered_line) >= 2:
            names.append(filtered_line)
            print(f"✅ ДОБАВЛЕНО ИМЯ: '{filtered_line}'")
        else:
            print(f"❌ СТРОКА ОТБРОШЕНА (пустая или слишком короткая)")
    
    return names

def add_names_to_excel(names_list):
    """
    Добавляет имена в Excel - каждый день в новый столбец
    """
    max_attempts = 5
    attempt = 0
    
    while attempt < max_attempts:
        try:
            # Текущая дата для заголовка столбца в формате ДД-ММ-ГГГГ
            today = datetime.now().strftime("%d-%m-%Y")
            
            # Читаем существующий файл или создаем новый
            if os.path.exists(EXCEL_FILE):
                # Читаем файл, явно указывая, что заголовки - строки
                df = pd.read_excel(EXCEL_FILE, dtype=str)
                print(f"Текущий DataFrame:\n{df}")
                print(f"Заголовки столбцов: {df.columns.tolist()}")
                
                # Сохраняем порядок столбцов
                original_columns_order = df.columns.tolist()
                print(f"Исходный порядок столбцов: {original_columns_order}")
                
                # Приводим все существующие заголовки к строковому формату ДД-ММ-ГГГГ
                new_columns = {}
                for col in df.columns:
                    col_str = str(col).strip()
                    
                    # Пробуем распарсить как дату в разных форматах
                    try:
                        # Проверяем формат ДД.ММ.ГГГГ
                        if re.match(r'^\d{2}\.\d{2}\.\d{4}$', col_str):
                            date_obj = datetime.strptime(col_str, "%d.%m.%Y")
                            new_col_name = date_obj.strftime("%d-%m-%Y")
                            new_columns[new_col_name] = df[col].tolist()
                        # Проверяем формат ГГГГ-ММ-ДД
                        elif re.match(r'^\d{4}-\d{2}-\d{2}', col_str):
                            # Извлекаем только дату без времени
                            date_part = col_str.split()[0] if ' ' in col_str else col_str
                            date_obj = datetime.strptime(date_part, "%Y-%m-%d")
                            new_col_name = date_obj.strftime("%d-%m-%Y")
                            new_columns[new_col_name] = df[col].tolist()
                        # Проверяем числовой формат (например 5.0, 5)
                        elif col_str.replace('.', '').isdigit():
                            # Пробуем интерпретировать как дату Excel
                            try:
                                # Excel даты - это числа, где 1 = 01.01.1900
                                excel_date = float(col_str)
                                if excel_date > 40000:  # Примерно 2009 год
                                    date_obj = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(excel_date) - 2)
                                    new_col_name = date_obj.strftime("%d-%m-%Y")
                                    new_columns[new_col_name] = df[col].tolist()
                                else:
                                    new_columns[col_str] = df[col].tolist()
                            except:
                                new_columns[col_str] = df[col].tolist()
                        else:
                            new_columns[col_str] = df[col].tolist()
                    except Exception as e:
                        print(f"Ошибка при обработке колонки {col_str}: {e}")
                        new_columns[col_str] = df[col].tolist()
                
                # Находим максимальную длину среди всех столбцов
                max_length = max(len(values) for values in new_columns.values())
                
                # Выравниваем все столбцы до максимальной длины
                for col_name in new_columns:
                    current_length = len(new_columns[col_name])
                    if current_length < max_length:
                        new_columns[col_name] = new_columns[col_name] + [None] * (max_length - current_length)
                
                # Создаем новый DataFrame с конвертированными заголовками
                df = pd.DataFrame(new_columns)
                
                # Восстанавливаем порядок столбцов (только те, что были)
                existing_columns = [col for col in original_columns_order if col in df.columns]
                other_columns = [col for col in df.columns if col not in original_columns_order]
                
                # Новый порядок: сначала существующие в исходном порядке, потом новые
                new_order = existing_columns + other_columns
                df = df[new_order]
                
                print(f"Заголовки после конвертации: {df.columns.tolist()}")
                print(f"Новый порядок столбцов: {new_order}")
                
                # Проверяем, есть ли уже столбец с сегодняшней датой
                if today in df.columns:
                    # Столбец с сегодняшней датой уже существует
                    print(f"Столбец {today} уже существует")
                    
                    # Получаем существующие имена за сегодня (убираем NaN)
                    existing_today_names = [x for x in df[today].tolist() if pd.notna(x)]
                    print(f"Существующие имена сегодня ({len(existing_today_names)}): {existing_today_names}")
                    print(f"Новые имена с фото ({len(names_list)}): {names_list}")
                    
                    # Только новые имена для сегодняшнего дня
                    new_names_today = []
                    for name in names_list:
                        if name not in existing_today_names:
                            new_names_today.append(name)
                    
                    print(f"Новые имена для добавления ({len(new_names_today)}): {new_names_today}")
                    
                    if new_names_today:
                        # Объединяем существующие и новые имена за сегодня
                        all_today_names = existing_today_names + new_names_today
                        
                        # Сохраняем все столбцы как списки
                        all_data = {}
                        for col in df.columns:
                            if col != today:
                                all_data[col] = df[col].tolist()
                        
                        # Добавляем обновленный today
                        all_data[today] = all_today_names
                        
                        # Находим максимальную длину
                        max_len = max(len(values) for values in all_data.values())
                        
                        # Выравниваем все столбцы
                        for col_name in all_data:
                            current_len = len(all_data[col_name])
                            if current_len < max_len:
                                all_data[col_name] = all_data[col_name] + [None] * (max_len - current_len)
                        
                        # Создаем новый DataFrame
                        new_df = pd.DataFrame(all_data)
                        
                        # Восстанавливаем порядок столбцов
                        new_df = new_df[df.columns.tolist()]
                        
                        df = new_df
                        added_count = len(new_names_today)
                        print(f"Добавлено {added_count} новых имен")
                    else:
                        added_count = 0
                        print("Новых имен нет")
                else:
                    # Сегодня еще не было записей - создаем новый столбец
                    print(f"Создаем новый столбец: {today}")
                    
                    # Удаляем дубликаты внутри новых имен
                    seen = set()
                    unique_new_names = []
                    for name in names_list:
                        if name not in seen:
                            seen.add(name)
                            unique_new_names.append(name)
                    
                    # Сохраняем все существующие столбцы как списки
                    all_data = {}
                    for col in df.columns:
                        all_data[col] = df[col].tolist()
                    
                    # Добавляем новый столбец
                    all_data[today] = unique_new_names
                    
                    # Находим максимальную длину
                    max_len = max(len(values) for values in all_data.values())
                    
                    # Выравниваем все столбцы
                    for col_name in all_data:
                        current_len = len(all_data[col_name])
                        if current_len < max_len:
                            all_data[col_name] = all_data[col_name] + [None] * (max_len - current_len)
                    
                    # Создаем новый DataFrame
                    new_df = pd.DataFrame(all_data)
                    
                    # Восстанавливаем порядок: существующие + новый в конец
                    new_order = df.columns.tolist() + [today]
                    df = new_df[new_order]
                    
                    added_count = len(unique_new_names)
                    print(f"Добавлено новых имен в новый столбец: {added_count}")
            else:
                # Новый файл - создаем с сегодняшней датой
                print(f"Создаем новый файл со столбцом: {today}")
                
                # Удаляем дубликаты из новых имен
                seen = set()
                unique_new_names = []
                for name in names_list:
                    if name not in seen:
                        seen.add(name)
                        unique_new_names.append(name)
                
                df = pd.DataFrame({
                    today: unique_new_names
                })
                
                added_count = len(unique_new_names)
                print(f"Добавлено новых имен в новый файл: {added_count}")
            
            # Сохраняем файл
            print(f"Сохраняем файл...")
            print(f"Итоговый порядок столбцов: {df.columns.tolist()}")
            df.to_excel(EXCEL_FILE, index=False)
            
            # Проверяем, сохранилось ли
            if os.path.exists(EXCEL_FILE):
                check_df = pd.read_excel(EXCEL_FILE, dtype=str)
                print(f"Проверка сохранения - прочитано:\n{check_df}")
                print(f"Заголовки столбцов: {check_df.columns.tolist()}")
                if today in check_df.columns:
                    saved_names = [x for x in check_df[today].tolist() if pd.notna(x)]
                    print(f"Количество записей в столбце {today}: {len(saved_names)}")
                    print(f"Сохраненные имена: {saved_names}")
            
            return added_count, today
            
        except PermissionError:
            attempt += 1
            if attempt < max_attempts:
                print(f"Файл заблокирован, попытка {attempt + 1} через 2 секунды")
                time.sleep(2)
            else:
                raise Exception("Файл Excel заблокирован. Закрой Excel и попробуй снова.")
        except Exception as e:
            print(f"Ошибка в add_names_to_excel: {e}")
            import traceback
            traceback.print_exc()
            raise e
        
def remove_duplicates_from_all_days():
    """
    Удаляет дубликаты ВНУТРИ каждого дня (столбца)
    """
    try:
        backup_path = make_backup()
        
        if not os.path.exists(EXCEL_FILE):
            return 0, None
        
        df = pd.read_excel(EXCEL_FILE)
        
        if df.empty:
            return 0, backup_path
        
        total_removed = 0
        
        # Для каждого столбца (дня) удаляем дубликаты
        for column in df.columns:
            # Получаем уникальные значения в столбце
            unique_values = df[column].dropna().tolist()
            original_count = len(unique_values)
            
            # Удаляем дубликаты
            seen = set()
            unique_ordered = []
            for value in unique_values:
                if value not in seen:
                    seen.add(value)
                    unique_ordered.append(value)
            
            removed_in_column = original_count - len(unique_ordered)
            total_removed += removed_in_column
            
            # Обновляем столбец
            df[column] = pd.Series(unique_ordered)
        
        if total_removed > 0:
            df.to_excel(EXCEL_FILE, index=False)
        
        return total_removed, backup_path
        
    except Exception as e:
        print(f"Ошибка: {e}")
        return 0, None

def get_table_stats():
    """Статистика таблицы"""
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
            if not df.empty:
                total_days = len(df.columns)
                total_names = df.count().sum()
                
                # Считаем уникальные имена по всем дням
                all_names = set()
                for column in df.columns:
                    all_names.update(df[column].dropna().tolist())
                
                return total_days, total_names, len(all_names)
        except Exception as e:
            print(f"Ошибка при чтении статистики: {e}")
    return 0, 0, 0

@bot.message_handler(commands=['start'])
def send_welcome(message):
    total_days, total_names, unique_names = get_table_stats()
    bot.reply_to(message, 
        f"👋 Привет! Я бот для упрощения жизни самой лучшей кисы\n\n"
        f"📸 Отправляй скрины участников конференции (можно несколько фото в одном сообщении!)\n\n"
        f"📊 Статистика:\n"
        f"• Всего дней: {total_days}\n"
        f"• Всего записей: {total_names}\n"
        f"• Уникальных имен: {unique_names}\n\n"
        f"Команды:\n"
        f"/debug - показывает содержимое актуальной таблицы\n"
        f"/remove_duplicates - удаляет дубликаты\n"
        f"/help - показывает инструкцию по использованию бота\n"
        f"/delete_today - удаляет столбец с сегодняшней датой"
    )

@bot.message_handler(commands=['delete_today'])
def handle_delete_today(message):
    """Удаляет столбец с сегодняшней датой"""
    try:
        # Текущая дата в формате ДД-ММ-ГГГГ
        today = datetime.now().strftime("%d-%m-%Y")
        
        processing_msg = bot.reply_to(message, f"🔄 Проверяю наличие столбца {today}...")
        
        if not os.path.exists(EXCEL_FILE):
            bot.edit_message_text(
                "❌ Файл не найден",
                chat_id=message.chat.id,
                message_id=processing_msg.message_id
            )
            return
        
        # Читаем файл
        df = pd.read_excel(EXCEL_FILE, dtype=str)
        
        if df.empty or len(df.columns) == 0:
            bot.edit_message_text(
                "❌ В файле нет данных",
                chat_id=message.chat.id,
                message_id=processing_msg.message_id
            )
            return
        
        # Проверяем, есть ли сегодняшний столбец
        if today not in df.columns:
            bot.edit_message_text(
                f"❌ Столбец с датой {today} не найден",
                chat_id=message.chat.id,
                message_id=processing_msg.message_id
            )
            return
        
        # Получаем количество записей в сегодняшнем столбце
        names_count = df[today].dropna().count()
        column_position = list(df.columns).index(today) + 1  # +1 для человеко-читаемого формата
        
        # Создаем клавиатуру для подтверждения
        from telebot import types
        markup = types.InlineKeyboardMarkup()
        btn_yes = types.InlineKeyboardButton("✅ Да, удалить", callback_data=f"confirm_delete_today_{today}")
        btn_no = types.InlineKeyboardButton("❌ Нет, отмена", callback_data="cancel_delete")
        markup.add(btn_yes, btn_no)
        
        bot.edit_message_text(
            f"⚠️ ВНИМАНИЕ!\n\n"
            f"Вы собираетесь удалить сегодняшний столбец:\n"
            f"📅 {today}\n"
            f"📍 Позиция: столбец {column_position}\n"
            f"📊 Записей в столбце: {names_count}\n\n"
            f"Это действие нельзя отменить без резервной копии.\n"
            f"Подтвердите удаление:",
            chat_id=message.chat.id,
            message_id=processing_msg.message_id,
            reply_markup=markup
        )
        
    except Exception as e:
        bot.reply_to(message, f"❌ Ошибка: {str(e)}")
        print(f"Ошибка в handle_delete_today: {e}")
        import traceback
        traceback.print_exc()

@bot.callback_query_handler(func=lambda call: True)
def handle_callback(call):
    """Обработчик нажатий на кнопки"""
    try:
        if call.data == "cancel_delete":
            bot.edit_message_text(
                "✅ Удаление отменено",
                chat_id=call.message.chat.id,
                message_id=call.message.message_id
            )
            bot.answer_callback_query(call.id, "Отменено")
            
        elif call.data.startswith("confirm_delete_today_"):
            today = call.data.replace("confirm_delete_today_", "")
            
            bot.edit_message_text(
                f"🔄 Удаляю столбец {today}...",
                chat_id=call.message.chat.id,
                message_id=call.message.message_id
            )
            
            # Создаем резервную копию перед удалением
            backup_path = make_backup()
            
            # Читаем файл
            df = pd.read_excel(EXCEL_FILE, dtype=str)
            
            # Удаляем сегодняшний столбец
            df = df.drop(columns=[today])
            
            # Сохраняем файл
            df.to_excel(EXCEL_FILE, index=False)
            
            # Получаем обновленную статистику
            total_days, total_names, unique_names = get_table_stats()
            
            # Формируем сообщение о результате
            backup_info = f"\n💾 Резервная копия: {backup_path}" if backup_path else "\n⚠️ Резервная копия не создана"
            
            bot.edit_message_text(
                f"✅ Столбец {today} успешно удален!\n\n"
                f"📊 Текущая статистика:\n"
                f"• Всего дней: {total_days}\n"
                f"• Всего записей: {total_names}\n"
                f"• Уникальных имен: {unique_names}{backup_info}",
                chat_id=call.message.chat.id,
                message_id=call.message.message_id
            )
            
            # Отправляем обновленный файл
            with open(EXCEL_FILE, 'rb') as file:
                bot.send_document(call.message.chat.id, file)
            
            bot.answer_callback_query(call.id, "Столбец удален")
            
    except Exception as e:
        bot.answer_callback_query(call.id, f"Ошибка: {str(e)}")
        bot.edit_message_text(
            f"❌ Ошибка при удалении: {str(e)}",
            chat_id=call.message.chat.id,
            message_id=call.message.message_id
        )
        print(f"Ошибка в handle_callback: {e}")
        import traceback
        traceback.print_exc()

# Команда для быстрого удаления без подтверждения (для разработчика)
@bot.message_handler(commands=['force_delete_today'])
def handle_force_delete_today(message):
    """Принудительно удаляет сегодняшний столбец без подтверждения (только для разработчика)"""
    try:
        today = datetime.now().strftime("%d-%m-%Y")
        
        processing_msg = bot.reply_to(message, f"🔄 Принудительно удаляю столбец {today}...")
        
        if not os.path.exists(EXCEL_FILE):
            bot.edit_message_text(
                "❌ Файл не найден",
                chat_id=message.chat.id,
                message_id=processing_msg.message_id
            )
            return
        
        df = pd.read_excel(EXCEL_FILE, dtype=str)
        
        if today not in df.columns:
            bot.edit_message_text(
                f"❌ Столбец {today} не найден",
                chat_id=message.chat.id,
                message_id=processing_msg.message_id
            )
            return
        
        # Создаем резервную копию
        backup_path = make_backup()
        
        # Удаляем столбец
        df = df.drop(columns=[today])
        df.to_excel(EXCEL_FILE, index=False)
        
        total_days, total_names, unique_names = get_table_stats()
        
        bot.edit_message_text(
            f"✅ Столбец {today} принудительно удален!\n\n"
            f"📊 Статистика:\n"
            f"• Всего дней: {total_days}\n"
            f"• Всего записей: {total_names}\n"
            f"• Уникальных имен: {unique_names}",
            chat_id=message.chat.id,
            message_id=processing_msg.message_id
        )
        
        with open(EXCEL_FILE, 'rb') as file:
            bot.send_document(message.chat.id, file)
            
    except Exception as e:
        bot.reply_to(message, f"❌ Ошибка: {str(e)}")

@bot.message_handler(commands=['help'])
def send_help(message):
    bot.reply_to(message,
        "🤖 Как работает бот:\n\n"
        "1. Отправляете фото со списком участников (можно несколько в одном сообщении)\n"
        "2. Бот распознает текст\n"
        "3. 🔥 ВАЖНО: удаляются слова короче 3 символов\n"
        "4. Удаляются цифры из строк\n"
        "5. 📅 КАЖДЫЙ ДЕНЬ - НОВЫЙ СТОЛБЕЦ\n"
        "6. Внутри одного дня дубликаты НЕ добавляются\n\n"
        "📋 Пример:\n"
        "• Первое фото сегодня: Иван Петров, Анна Смирнова → столбец A\n"
        "• Второе фото сегодня: Иван Петров, Петр Иванов → добавится только Петр Иванов в столбец A\n"
        "• Завтра: Иван Петров → новый столбец B\n\n"
        "Команды:\n"
        "/help - показывает инструкцию по использованию бота\n"
        "/debug - показывает содержимое актуальной таблицы\n"
        "/remove_duplicates - удаляет дубликаты\n"
        "/delete_today - удаляет столбец с сегодняшней датой"
    )

@bot.message_handler(commands=['debug'])
def handle_debug(message):
    """Команда для отладки - показывает содержимое файла"""
    try:
        if os.path.exists(EXCEL_FILE):
            df = pd.read_excel(EXCEL_FILE)
            debug_msg = "📊 Содержимое файла:\n\n"
            for col in df.columns:
                names = df[col].dropna().tolist()
                debug_msg += f"📅 {col} ({len(names)} записей):\n"
                if names:
                    debug_msg += f"{', '.join(names[:5])}"
                    if len(names) > 5:
                        debug_msg += f" и еще {len(names) - 5}"
                else:
                    debug_msg += "Нет записей"
                debug_msg += "\n\n"
            bot.reply_to(message, debug_msg)
        else:
            bot.reply_to(message, "❌ Файл еще не создан")
    except Exception as e:
        bot.reply_to(message, f"❌ Ошибка: {e}")

@bot.message_handler(commands=['remove_duplicates'])
def handle_remove_duplicates(message):
    """Команда для удаления дубликатов внутри каждого дня"""
    try:
        processing_msg = bot.reply_to(message, "🔄 Удаляю дубликаты внутри каждого дня...")
        
        removed_count, backup_path = remove_duplicates_from_all_days()
        
        if removed_count > 0:
            total_days, total_names, unique_names = get_table_stats()
            bot.edit_message_text(
                f"✅ Дубликаты удалены!\n\n"
                f"🗑 Удалено записей: {removed_count}\n"
                f"📊 Текущая статистика:\n"
                f"• Всего дней: {total_days}\n"
                f"• Всего записей: {total_names}\n"
                f"• Уникальных имен: {unique_names}",
                chat_id=message.chat.id,
                message_id=processing_msg.message_id
            )
            
            # Отправляем файл
            with open(EXCEL_FILE, 'rb') as file:
                bot.send_document(message.chat.id, file)
        else:
            bot.edit_message_text(
                "ℹ️ Дубликатов не найдено или файл пуст",
                chat_id=message.chat.id,
                message_id=processing_msg.message_id
            )
            
    except Exception as e:
        bot.reply_to(message, f"❌ Ошибка: {str(e)}")

# ========== ЕДИНСТВЕННЫЙ ОБРАБОТЧИК ФОТО ==========
@bot.message_handler(content_types=['photo'])
def handle_photo(message):
    try:
        # Проверяем, является ли фото частью медиа-группы (альбома)
        if message.media_group_id:
            group_id = message.media_group_id
            
            # Если это первое фото в группе, создаем запись
            if group_id not in media_groups:
                media_groups[group_id] = {
                    'photos': [],
                    'message': None,
                    'chat_id': message.chat.id,
                    'processing': False,
                    'timer_started': False
                }
            
            # Добавляем фото в группу
            media_groups[group_id]['photos'].append(message)
            photos_count = len(media_groups[group_id]['photos'])
            
            # Если это первое фото, запускаем таймер
            if not media_groups[group_id]['timer_started']:
                media_groups[group_id]['timer_started'] = True
                
                timer = threading.Timer(2.0, process_media_group, args=[group_id])
                timer.start()
                
                # Отправляем подтверждение
                msg = bot.reply_to(message, f"📸 Получено фото 1. Ожидание остальных...")
                media_groups[group_id]['message'] = msg
            else:
                # Обновляем сообщение о количестве полученных фото
                if media_groups[group_id]['message']:
                    try:
                        bot.edit_message_text(
                            f"📸 Получено фото {photos_count}. Ожидание остальных...",
                            chat_id=message.chat.id,
                            message_id=media_groups[group_id]['message'].message_id
                        )
                    except:
                        pass  # Игнорируем ошибки редактирования
        else:
            # Одиночное фото - обрабатываем как обычно
            process_single_photo(message)
            
    except Exception as e:
        bot.reply_to(message, f"❌ Ошибка: {str(e)}")
        print(f"Ошибка в handle_photo: {e}")
        import traceback
        traceback.print_exc()

def process_media_group(group_id):
    """Обрабатывает все фото из медиа-группы"""
    try:
        if group_id not in media_groups:
            return
        
        group_data = media_groups[group_id]
        
        # Проверяем, не обрабатывается ли уже эта группа
        if group_data['processing']:
            return
        
        group_data['processing'] = True
        photos = group_data['photos']
        chat_id = group_data['chat_id']
        status_msg = group_data['message']
        
        # Обновляем статус
        if status_msg:
            try:
                bot.edit_message_text(
                    f"🔄 Обрабатываю {len(photos)} фото...",
                    chat_id=chat_id,
                    message_id=status_msg.message_id
                )
            except:
                pass
        
        # Список для всех имен с каждого фото
        all_names = []
        photos_processed = 0
        photos_failed = 0
        
        for i, photo_msg in enumerate(photos, 1):
            try:
                # Обновляем статус
                if status_msg:
                    try:
                        bot.edit_message_text(
                            f"🔄 Обрабатываю фото {i}/{len(photos)}...",
                            chat_id=chat_id,
                            message_id=status_msg.message_id
                        )
                    except:
                        pass
                
                # Получаем фото
                file_info = bot.get_file(photo_msg.photo[-1].file_id)
                downloaded_file = bot.download_file(file_info.file_path)
                
                # Сохраняем
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                photo_path = f"temp/photo_{timestamp}_{i}.jpg"
                
                with open(photo_path, 'wb') as new_file:
                    new_file.write(downloaded_file)
                
                # Извлекаем имена
                names = extract_names_from_image(photo_path)
                
                # Удаляем фото
                try:
                    os.remove(photo_path)
                except:
                    pass
                
                if names:
                    all_names.extend(names)
                    photos_processed += 1
                    print(f"Фото {i}: найдено {len(names)} имен: {names}")
                else:
                    photos_failed += 1
                    print(f"Фото {i}: имен не найдено")
                    
            except Exception as e:
                photos_failed += 1
                print(f"Ошибка при обработке фото {i}: {e}")
        
        if not all_names:
            if status_msg:
                try:
                    bot.edit_message_text(
                        f"❌ Не удалось найти имена ни на одном из {len(photos)} фото",
                        chat_id=chat_id,
                        message_id=status_msg.message_id
                    )
                except:
                    pass
            return
        
        # Показываем что нашли
        unique_names = list(set(all_names))  # Временная уникализация для предпросмотра
        names_text = "\n".join([f"• {name}" for name in unique_names[:10]])
        if len(unique_names) > 10:
            names_text += f"\n• ... и еще {len(unique_names) - 10}"
        
        if status_msg:
            try:
                bot.edit_message_text(
                    f"✅ Найдено всего: {len(unique_names)} имен с {photos_processed} фото\n"
                    f"❌ Без имен: {photos_failed} фото\n\n"
                    f"{names_text}\n\n📝 Добавляю в Excel...",
                    chat_id=chat_id,
                    message_id=status_msg.message_id
                )
            except:
                pass
        
        # Добавляем в Excel
        added_count, today_date = add_names_to_excel(all_names)
        
        # Получаем обновленную статистику
        total_days, total_names, unique_all_names = get_table_stats()
        
        if added_count == 0:
            final_msg = f"ℹ️ За {today_date} все имена уже были добавлены ранее\n\n"
        else:
            final_msg = f"✅ Готово!\n\n📅 {today_date}\n➕ Новых: {added_count}\n🔄 Повторов: {len(all_names) - added_count}\n\n"
        
        final_msg += f"📊 Статистика:\n• Всего дней: {total_days}\n• Всего записей: {total_names}\n• Уникальных имен: {unique_all_names}"
        
        if status_msg:
            try:
                bot.edit_message_text(
                    final_msg,
                    chat_id=chat_id,
                    message_id=status_msg.message_id
                )
            except:
                pass
        
        # Отправляем файл
        with open(EXCEL_FILE, 'rb') as file:
            bot.send_document(chat_id, file)
        
    except Exception as e:
        print(f"Ошибка в process_media_group: {e}")
        import traceback
        traceback.print_exc()
        if 'status_msg' in locals() and status_msg:
            try:
                bot.edit_message_text(
                    f"❌ Ошибка при обработке группы: {str(e)}",
                    chat_id=chat_id,
                    message_id=status_msg.message_id
                )
            except:
                pass
    finally:
        # Очищаем данные группы
        if group_id in media_groups:
            del media_groups[group_id]

def process_single_photo(message):
    """Обрабатывает одиночное фото"""
    try:
        processing_msg = bot.reply_to(message, "🔄 Обрабатываю фото...")
        
        # Получаем фото
        file_info = bot.get_file(message.photo[-1].file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        # Сохраняем
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        photo_path = f"temp/photo_{timestamp}.jpg"
        
        with open(photo_path, 'wb') as new_file:
            new_file.write(downloaded_file)
        
        bot.edit_message_text(
            "🔍 Сканирую текст...", 
            chat_id=message.chat.id, 
            message_id=processing_msg.message_id
        )
        
        # Извлекаем имена
        names = extract_names_from_image(photo_path)
        
        # Удаляем фото
        try:
            os.remove(photo_path)
        except:
            pass
        
        if not names:
            bot.edit_message_text(
                "❌ Не удалось найти имена на фото",
                chat_id=message.chat.id,
                message_id=processing_msg.message_id
            )
            return
        
        # Показываем что нашли
        names_text = "\n".join([f"• {name}" for name in names[:10]])
        if len(names) > 10:
            names_text += f"\n• ... и еще {len(names) - 10}"
        
        bot.edit_message_text(
            f"✅ Найдено: {len(names)}\n\n{names_text}\n\n📝 Добавляю в Excel...",
            chat_id=message.chat.id,
            message_id=processing_msg.message_id
        )
        
        # Добавляем в Excel
        added_count, today_date = add_names_to_excel(names)
        
        # Получаем обновленную статистику
        total_days, total_names, unique_names = get_table_stats()
        
        if added_count == 0:
            bot.edit_message_text(
                f"ℹ️ За {today_date} все имена уже были добавлены ранее\n\n"
                f"📊 Статистика:\n"
                f"• Всего дней: {total_days}\n"
                f"• Всего записей: {total_names}\n"
                f"• Уникальных имен: {unique_names}",
                chat_id=message.chat.id,
                message_id=processing_msg.message_id
            )
        else:
            bot.edit_message_text(
                f"✅ Готово!\n\n"
                f"📅 {today_date}\n"
                f"➕ Новых: {added_count}\n"
                f"🔄 Повторов: {len(names) - added_count}\n\n"
                f"📊 Статистика:\n"
                f"• Всего дней: {total_days}\n"
                f"• Всего записей: {total_names}\n"
                f"• Уникальных имен: {unique_names}\n\n"
                f"📎 Отправляю файл...",
                chat_id=message.chat.id,
                message_id=processing_msg.message_id
            )
            
            # Отправляем файл
            with open(EXCEL_FILE, 'rb') as file:
                bot.send_document(message.chat.id, file)
        
    except Exception as e:
        bot.reply_to(message, f"❌ Ошибка: {str(e)}")
        print(f"Ошибка в process_single_photo: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    print("🚀 Бот запущен...")
    print(f"📁 Файл: {EXCEL_FILE}")
    print(f"• Сегодня ({datetime.now().strftime('%d.%m.%Y')})")
    print("📸 Поддержка нескольких фото в одном сообщении!")
    print("📝 Фильтр: слова от 3 букв, цифры удаляются")
    print("💡 Для отладки используй /debug")
    print("   Для удаления дубликатов используй /remove_duplicates")
    bot.infinity_polling()