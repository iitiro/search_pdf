import requests
import pandas as pd
import os
import json
from datetime import datetime

# Шляхи до файлів з API ключем та ідентифікатором пошукової системи
api_key_file = '/Users/ikudinov/Documents/Code/keys/api.txt'
engine_id_file = '/Users/ikudinov/Documents/Code/keys/engine.txt'

# Перевірка наявності файлу з API ключем
if not os.path.exists(api_key_file):
    print(f"Файл {api_key_file} не знайдено.")
    exit(1)

# Перевірка наявності файлу з engine ID
if not os.path.exists(engine_id_file):
    print(f"Файл {engine_id_file} не знайдено.")
    exit(1)

# Читання API ключа з файлу
with open(api_key_file, 'r') as file:
    api_key = file.read().strip()

# Читання ідентифікатора пошукової системи з файлу
with open(engine_id_file, 'r') as file:
    cx = file.read().strip()

# Читаємо ключові слова з файлу "keywords.txt"
keywords_file = 'keywords.txt'

# Перевіряємо, чи існує файл з ключовими словами
if not os.path.exists(keywords_file):
    print(f"Файл {keywords_file} не знайдено.")
    exit(1)

with open(keywords_file, 'r') as file:
    keywords = [line.strip() for line in file if line.strip()]

# Максимальна кількість файлів для кожного ключового слова
max_results = 40

# Основна папка для збереження всіх результатів
main_folder = '!search_pdf'
os.makedirs(main_folder, exist_ok=True)

# Поточна дата і час для позначення запиту
current_time = datetime.now().strftime('%Y-%m-%d %H-%M')

# Проходимося по кожному ключовому слову
for keyword in keywords:
    # Пропускаємо ключові слова, які починаються з символу #
    if keyword.startswith("#"):
        print(f"Пропуск ключового слова: {keyword}")
        continue

    print(f"Обробка ключового слова: {keyword}")

    # Створюємо підпапку для ключового слова з додаванням дати і часу запиту
    keyword_folder = os.path.join(main_folder, f"{current_time} {keyword}")
    os.makedirs(keyword_folder, exist_ok=True)

    # Список для зберігання даних
    data = []
    start_index = 1

    while len(data) < max_results:
        # Формуємо запит до Google Custom Search API
        query = f'"{keyword}" filetype:pdf'
        url = f"https://www.googleapis.com/customsearch/v1?q={query}&key={api_key}&cx={cx}&start={start_index}"

        # Відправлення запиту до API
        response = requests.get(url)
        results = response.json()

        # Перевірка на наявність результатів
        if 'items' in results:
            # Обробка кожного результату
            for i, item in enumerate(results['items']):
                title = item.get('title', 'No title')
                link = item.get('link', 'No link')
                snippet = item.get('snippet', 'No description')
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Час для кожного запиту
                status = 'Not Attempted'  # Статус за замовчуванням

                # Якщо посилання веде до PDF
                if link.endswith('.pdf'):
                    # Спробуємо скачати PDF файл
                    try:
                        pdf_response = requests.get(link, timeout=10)  # Вказуємо тайм-аут
                        pdf_response.raise_for_status()  # Перевіряємо, чи немає помилки в відповіді

                        # Зберігаємо PDF-файл в підпапку ключового слова
                        pdf_name = os.path.join(keyword_folder, f"file_{len(data) + 1}.pdf")
                        with open(pdf_name, 'wb') as f:
                            f.write(pdf_response.content)

                        status = 'Успішно завантажено'
                        print(f"Файл {pdf_name} завантажено успішно.")

                    except requests.exceptions.Timeout:
                        status = 'Тайм-аут при завантаженні'
                        print(f"Тайм-аут при завантаженні файлу: {link}")
                    except requests.exceptions.RequestException as e:
                        status = f"Помилка при завантаженні: {e}"
                        print(f"Помилка при завантаженні файлу {link}: {e}")

                # Додаємо дані в таблицю
                data.append([len(data) + 1, title, snippet, link, status, timestamp])

                # Перевірка ліміту
                if len(data) >= max_results:
                    break

            # Оновлення start_index для наступного запиту
            start_index += 10
        else:
            # Якщо результів немає, цикл завершується
            print("Немає більше результатів пошуку.")
            break

    # Трансформація даних в DataFrame
    df = pd.DataFrame(data, columns=['Index', 'Title', 'Description', 'Link', 'Status', 'Timestamp'])

    # Збереження результатів в Excel з додаванням дати і часу до назви файлу
    excel_file_path = os.path.join(keyword_folder, f'pdf_files_data_{current_time}.xlsx')
    df.to_excel(excel_file_path, index=False, engine='openpyxl')

    print(f"Готово! Збережено {len(data)} файлів і даних для ключового слова '{keyword}'.")

print("Обробка всіх ключових слів завершена.")