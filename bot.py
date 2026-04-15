import telebot
import os
import base64
import json
import pandas as pd
from datetime import datetime
from openai import OpenAI
import shutil
import time
import re

# Токен бота
TOKEN = '8697524461:AAFs8il54OBoGjs8VnrvoQGkgplvxuYUDZ8'
# Ключ OpenAI API (для Vision / chat completions)
OPENAI_API_KEY = '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'
OPENAI_VISION_MODEL = 'gpt-4o'

bot = telebot.TeleBot(TOKEN)

# Создаем папки
for folder in ['temp', 'backups']:
    if not os.path.exists(folder):
        os.makedirs(folder)

EXCEL_FILE = 'ебанаты (кроме кисули).xlsx'

# Слова для пропуска строк
SKIP_WORDS = ['участник', 'поиск', 'организатор', 'конференции', 'участники']

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

def filter_short_words(text):
    """
    Фильтрует слова, оставляя только те, что длиннее 3 символов (минимум 4 символа)
    """
    if not text:
        return text
    
    # Разбиваем на слова
    words = text.split()
    
    # Фильтруем слова (оставляем только те, где > 3 символов, т.е. минимум 4)
    filtered_words = [word for word in words if len(word) > 3]
    
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
    for word in SKIP_WORDS:
        if word in line_lower:
            return True
    return False

def _strip_json_fence(text):
    text = text.strip()
    if text.startswith("```"):
        lines = text.split("\n")
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        text = "\n".join(lines)
    return text.strip()

def _parse_gpt_names_json(raw):
    raw = _strip_json_fence(raw)
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        print(f"Не удалось разобрать JSON от модели: {raw[:500]}")
        return []
    if isinstance(data, dict) and "names" in data:
        data = data["names"]
    if not isinstance(data, list):
        return []
    out = []
    for item in data:
        if isinstance(item, str) and item.strip():
            out.append(item.strip())
    return out

def _image_mime_for_path(path):
    ext = os.path.splitext(path)[1].lower()
    return {
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".gif": "image/gif",
        ".webp": "image/webp",
    }.get(ext, "image/jpeg")

def extract_names_from_image(image_path):
    """
    Извлекает имена с фото через OpenAI Vision.
    """
    if not OPENAI_API_KEY:
        raise RuntimeError("Заполните OPENAI_API_KEY в начале bot.py")

    with open(image_path, "rb") as f:
        b64 = base64.standard_b64encode(f.read()).decode("ascii")

    mime = _image_mime_for_path(image_path)

    client = OpenAI(api_key=OPENAI_API_KEY)
    response = client.chat.completions.create(
        model=OPENAI_VISION_MODEL,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": (
                            "На изображении список людей (часто участники конференции или поиска). "
                            "Извлеки полные имена или ФИО: одна строка — один человек. "
                            "Не включай заголовки, кнопки интерфейса, служебный текст, URL, подписи вроде «участник». "
                            "Верни строго JSON-массив строк в UTF-8, без пояснений до или после, "
                            'например: ["Иван Петров","Мария Иванова"]. Если подходящих имён нет — [].'
                        ),
                    },
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:{mime};base64,{b64}"},
                    },
                ],
            }
        ],
        max_tokens=4096,
    )

    raw = (response.choices[0].message.content or "").strip()
    lines = _parse_gpt_names_json(raw)

    names = []
    for line in lines:
        line = clean_text(line)
        if not line:
            continue
        if should_skip_line(line):
            continue
        formatted_line = capitalize_words(line)
        formatted_line = filter_short_words(formatted_line)
        formatted_line = clean_text(formatted_line)
        if formatted_line and len(formatted_line) >= 2:
            names.append(formatted_line)

    return names

def add_names_to_excel(names_list):
    """
    Добавляет имена в Excel - каждый день в новый столбец
    """
    max_attempts = 5
    attempt = 0
    
    while attempt < max_attempts:
        try:
            # Текущая дата для заголовка столбца
            today = datetime.now().strftime("%d.%m.%Y")
            
            # Читаем существующий файл или создаем новый
            if os.path.exists(EXCEL_FILE):
                df = pd.read_excel(EXCEL_FILE)
                print(f"Текущий DataFrame:\n{df}")
                print(f"Индекс DataFrame: {df.index.tolist()}")
                
                # Проверяем, есть ли уже столбец с сегодняшней датой
                if today in df.columns:
                    # Столбец с сегодняшней датой уже существует
                    print(f"Столбец {today} уже существует")
                    
                    # Получаем существующие имена за сегодня (убираем NaN)
                    existing_today_names = df[today].dropna().tolist()
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
                        
                        # Создаем новый DataFrame с правильной структурой
                        # Сохраняем все остальные столбцы
                        other_columns = {}
                        for col in df.columns:
                            if col != today:
                                other_columns[col] = df[col].tolist()
                        
                        # Создаем новый DataFrame с обновленным столбцом
                        new_df = pd.DataFrame({
                            today: all_today_names
                        })
                        
                        # Добавляем остальные столбцы
                        for col_name, col_values in other_columns.items():
                            # Приводим к одинаковой длине
                            max_len = max(len(new_df), len(col_values))
                            padded_values = col_values + [None] * (max_len - len(col_values))
                            new_df[col_name] = padded_values[:max_len]
                        
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
                    
                    # Приводим все столбцы к одинаковой длине
                    max_len = max(len(df), len(unique_new_names))
                    
                    # Создаем новый DataFrame с правильной длиной
                    new_df = pd.DataFrame()
                    
                    # Копируем существующие столбцы с паддингом
                    for col in df.columns:
                        col_values = df[col].tolist()
                        padded_values = col_values + [None] * (max_len - len(col_values))
                        new_df[col] = padded_values[:max_len]
                    
                    # Добавляем новый столбец
                    new_column = unique_new_names + [None] * (max_len - len(unique_new_names))
                    new_df[today] = new_column
                    
                    df = new_df
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
            print(f"Итоговый DataFrame:\n{df}")
            df.to_excel(EXCEL_FILE, index=False)
            
            # Проверяем, сохранилось ли
            if os.path.exists(EXCEL_FILE):
                check_df = pd.read_excel(EXCEL_FILE)
                print(f"Проверка сохранения - прочитано:\n{check_df}")
                if today in check_df.columns:
                    saved_names = check_df[today].dropna().tolist()
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
        f"📸 Отправляй скрины участников конференции по одному в сообщении, дожидаясь формирования таблицы\n\n"
        f"📊 Статистика:\n"
        f"• Всего дней: {total_days}\n"
        f"• Всего записей: {total_names}\n"
        f"• Уникальных имен: {unique_names}\n\n"
        f"Команды:\n"
        f"/debug - показывает содержимое актуальной таблицы\n"
        f"/remove_duplicates - удаляет дубликаты\n"
        f"/help - показывает инструкцию по использованию бота"
    )

@bot.message_handler(commands=['help'])
def send_help(message):
    bot.reply_to(message,
        "🤖 Как работает бот:\n\n"
        "1. Отправляете фото со списком участников\n"
        "2. Бот отправляет фото в OpenAI (GPT-4 Vision) и получает список имён\n"
        "3. 🔥 ВАЖНО: удаляются слова короче 4 символов\n"
        "4. 📅 КАЖДЫЙ ДЕНЬ - НОВЫЙ СТОЛБЕЦ\n"
        "5. Внутри одного дня дубликаты НЕ добавляются\n\n"
        "📋 Пример:\n"
        "• Первое фото сегодня: Иван Петров, Анна Смирнова → столбец A\n"
        "• Второе фото сегодня: Иван Петров, Петр Иванов → добавится только Петр Иванов в столбец A\n"
        "• Завтра: Иван Петров → новый столбец B\n\n"
        "Команды:\n"
        "/debug - показывает содержимое актуальной таблицы\n"
        "/remove_duplicates - удаляет дубликаты\n"
        "/help - показывает инструкцию по использованию бота\n\n"
        "позже реализую команду, очищающую столбец по указанной дате"
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

@bot.message_handler(content_types=['photo'])
def handle_photo(message):
    try:
        if message.media_group_id:
            bot.reply_to(message, "❌ Отправляй только ОДНО фото")
            return
        
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
            "🤖 Анализ изображения (GPT-4 Vision)...",
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
        
        # Добавляем в Excel (каждый день в новый столбец)
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
                f"📅 {today_date} (столбец {chr(64 + total_days) if total_days <= 26 else 'XX'})\n"
                f"➕ Новых за сегодня: {added_count}\n"
                f"🔄 Повторов сегодня: {len(names) - added_count}\n\n"
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
        print(f"Ошибка в handle_photo: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    print("🚀 Бот запущен...")
    if not OPENAI_API_KEY:
        print("⚠️  Пустой OPENAI_API_KEY в bot.py — обработка фото не будет работать.")
    else:
        print(f"🔑 OpenAI Vision: {OPENAI_VISION_MODEL}")
    print(f"📁 Файл: {EXCEL_FILE}")
    print(f"•  Сегодня ({datetime.now().strftime('%d.%m.%Y')})")
    print("💡 Для отладки используй /debug")
    print("   Для удаления дубликатов используй /remove_duplicates")
    bot.infinity_polling()