import rpa as r
import pandas as pd
import time
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

import PyPDF2
import docx
import unicodedata
import codecs

# Функция для нормализации имени файла (для работы с русскими символами)
def normalize_filename(name):
    return unicodedata.normalize('NFKD', name)

# Функция для поиска скачанного файла с учётом нормализации имени
def find_downloaded_file(expected_name, download_folder, timeout=60):
    normalized_expected = normalize_filename(expected_name)
    start_time = time.time()
    while time.time() - start_time < timeout:
        files = os.listdir(download_folder)
        for f in files:
            normalized_f = normalize_filename(f)
            # Если нормализованное ожидаемое имя встречается в нормализованном имени файла, возвращаем полный путь
            if normalized_expected in normalized_f:
                return os.path.join(download_folder, f)
        time.sleep(1)
    return None

# Функция для поиска искомого слова в PDF-файле
def search_in_pdf(filepath, search_word):
    try:
        with open(filepath, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text = page.extract_text()
                if text and search_word.lower() in text.lower():
                    return True
    except Exception as e:
        print(f"Ошибка чтения PDF {filepath}: {e}")
    return False

# Функция для поиска искомого слова в DOCX-файле
def search_in_docx(filepath, search_word):
    try:
        doc = docx.Document(filepath)
        for para in doc.paragraphs:
            if search_word.lower() in para.text.lower():
                return True
    except Exception as e:
        print(f"Ошибка чтения DOCX {filepath}: {e}")
    return False

# Функция для отправки письма с прикрепленным файлом мониторинга
def send_email(subject, body, attachment_path, recipient_email):
    # Пример SMTP-настроек; замените на реальные данные
    smtp_server = 'smtp.example.com'
    smtp_port = 587
    smtp_user = 'your_email@example.com'
    smtp_password = 'your_password'
    
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = recipient_email
    msg['Subject'] = subject
    
    msg.attach(MIMEText(body, 'plain'))
    
    with open(attachment_path, 'rb') as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
    msg.attach(part)
    
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(smtp_user, smtp_password)
    server.send_message(msg)
    server.quit()

# 1. Чтение файла с тегами (слова для поиска) из "Теги.xlsx"
df_tags = pd.read_excel("Теги.xlsx")
# Предполагается, что искомые слова находятся в первой колонке
search_words = df_tags.iloc[:, 0].dropna().astype(str).tolist()

# 2. Подготовка файла мониторинга (если он уже существует, загружаем его; иначе создаём новый)
monitoring_file = "Мониторинг.xlsx"
if os.path.exists(monitoring_file):
    monitoring_df = pd.read_excel(monitoring_file)
else:
    monitoring_df = pd.DataFrame(columns=["Ссылка на закупку", "Лот №", "Организатор", "Сумма закупки"])

# 3. Файл для хранения обработанных объявлений (по уникальному номеру объявления)
processed_file = "processed_announcements.txt"
if os.path.exists(processed_file):
    with open(processed_file, 'r') as f:
        processed_announcements = set(line.strip() for line in f if line.strip())
else:
    processed_announcements = set()

# 4. Инициализация RPA-сессии и открытие страницы поиска объявлений
r.init(visual_automation=True)
r.url("https://goszakup.gov.kz/ru/search/announce")
time.sleep(3)

# Основной цикл по каждому слову из Теги.xlsx
for word in search_words:
    print(f"\nОбрабатываем слово: {word}")
    # Для каждого слова запускаем поиск сначала по финансовому году 2020, затем по 2021
    for year in ["2020", "2021"]:
        print(f"\nПоиск для финансового года: {year}")
        # Вместо r.reload() переходим по URL страницы поиска для сброса формы
        r.url("https://goszakup.gov.kz/ru/search/announce")
        time.sleep(3)
        
        # Заполняем форму поиска
        print("Вводим наименование объявления...")
        r.type("//*[@id='in_name']", word)
        
        print("Выбираем способ закупки...")
        r.click("//*[@data-id='s2_method']")
        r.click("//*[text()='Запрос ценовых предложений']")
        r.click("//*[text()='Открытый конкурс']")
        
        print("Выбираем статус 'Завершено'...")
        r.click("//*[@data-id='s2_status']")
        r.click("//*[text()='Завершено']")
        
        print("Устанавливаем сумму закупки...")
        r.type("//*[@id='in_amount_from']", "1000000")
        
        print(f"Выбираем финансовый год {year}...")
        r.click("//*[@data-id='s2_year']")
        r.click(f"//*[text()='{year}']")
        
        print("Нажимаем кнопку 'Найти'...")
        r.click("//button[@name='smb' and contains(@class, 'btn-success')]")
        r.wait(5)
        
        # Перебор объявлений на текущей странице
        announcement_index = 1
        while True:
            announcement_xpath = f"//*[@id='search-result']//tr[{announcement_index}]//a"
            if not r.exist(announcement_xpath):
                print("Объявлений на странице больше нет.")
                break  # Выход из цикла, если объявлений больше нет
            
            # Извлекаем ссылку и определяем уникальный номер объявления
            href = r.read(announcement_xpath + "/@href")
            try:
                unique_id = href.split('/')[-1]
            except Exception:
                unique_id = None
            
            if unique_id in processed_announcements:
                print(f"Объявление {unique_id} уже обработано, пропускаем.")
                announcement_index += 1
                continue
            
            back_url = r.url()

            # //*[@id="search-result"]/tbody/tr[1]/td[2]/a
            # //*[@id="search-result"]/tbody/tr[2]/td[2]/a
            # //*[@id="search-result"]/tbody/tr[3]/td[2]/a

            print(announcement_xpath)
            print(f"\nОбрабатываем объявление {unique_id}...")
            r.click(f"//*[@id='search-result']//tr[{announcement_index}]//a")
            print(f"//*[@id='search-result']//tr[{announcement_index}]//a")
            time.sleep(3)
            # Читаем ссылку объявления и формируем полный URL
            announcement_url = r.read(announcement_xpath + "/@href")
            full_url = f"https://goszakup.gov.kz{announcement_url}" if announcement_url.startswith("/") else announcement_url
            # Открываем объявление через r.popup(), чтобы перевести фокус на новое окно
            r.popup(full_url)
            time.sleep(3)
            
            current_url = r.url()  # URL текущей вкладки (объявления)
            print("Обрабатываем страницу объявления:", current_url)
            
            # Переходим во вкладку "Документация"
            if r.exist("//*[text()='Документация']"):
                r.click("//*[text()='Документация']")
                time.sleep(2)
            else:
                print("Вкладка 'Документация' не найдена, пропускаем объявление.")
                processed_announcements.add(unique_id)
                # Закрываем всплывающее окно с объявлением
                r.keyboard("[ctrl]w")
                time.sleep(3)
                announcement_index += 1
                continue
            
            # Нажимаем кнопку "Перейти" для технической спецификации
            tech_spec_xpath = "//*[contains(text(), 'Техническая спецификация')]/following-sibling::td/button[contains(text(), 'Перейти')]"
            if r.exist(tech_spec_xpath):
                r.click(tech_spec_xpath)
                time.sleep(3)  # Ожидаем появления модального окна
            else:
                print("Кнопка 'Перейти' для Технической спецификации не найдена, пропускаем объявление.")
                processed_announcements.add(unique_id)
                # Закрываем всплывающее окно с объявлением
                r.keyboard("[ctrl]w")
                time.sleep(3)
                announcement_index += 1
                continue
            
            # В модальном окне со списком файлов перебираем строки таблицы (первая строка – заголовок)
            file_row_index = 2  # Начинаем со 2-й строки
            found_in_files = False
            file_found_lot = None
            while True:
                file_row_xpath = f"//*[@id='ModalShowFilesBody']/table/tbody/tr[{file_row_index}]"
                if not r.exist(file_row_xpath):
                    break  # Больше строк нет
                # Из первой колонки получаем номер лота
                lot_xpath = file_row_xpath + "/td[1]"
                lot_number = r.read(lot_xpath).strip()
                #lot_number = codecs.decode(lot_number, 'unicode-escape')
                # Из второй колонки получаем ссылку и имя файла
                file_link_xpath = file_row_xpath + "/td[2]/a"
                file_url = r.read(file_link_xpath + "/@href").strip()
                file_name = r.read(file_link_xpath).strip()
                #file_name = codecs.decode(file_name, 'unicode_escape')
                print(f"Обрабатываем файл: {file_name} (Лот: {lot_number})")
                
                # Кликаем по ссылке для скачивания файла
                r.click(f"//*[text()='{file_name}']")
                print(file_link_xpath)
                # Ожидаем, что файл скачан в текущую папку с учётом нормализации имени
                download_path = find_downloaded_file(file_name, os.getcwd(), timeout=60)
                if download_path is None:
                    print(f"Файл {file_name} не скачался вовремя.")
                    file_row_index += 1
                    continue
                else:
                    print(f"Найден скачанный файл: {download_path}")
                
                # Поиск искомого слова в файле (в зависимости от расширения)
                found = False
                if file_name.lower().endswith('.pdf'):
                    found = search_in_pdf(download_path, word)
                elif file_name.lower().endswith('.docx'):
                    found = search_in_docx(download_path, word)
                else:
                    print(f"Неподдерживаемый формат файла: {file_name}")
                
                if found:
                    print(f"Слово '{word}' найдено в файле {file_name}.")
                    found_in_files = True
                    file_found_lot = lot_number
                    break  # Прекращаем перебор файлов для этого объявления
                else:
                    print(f"Слово '{word}' не найдено в файле {file_name}.")
                    file_row_index += 1
            
            # Закрываем модальное окно (если присутствует кнопка "Закрыть")
            if r.exist("//button[@data-dismiss='modal']"):
                r.click("//button[@data-dismiss='modal']")
                time.sleep(2)
            
            # Если слово найдено хотя бы в одном файле, переходим во вкладку "Общие сведения"
            if found_in_files:
                if r.exist("//*[text()='Общие сведения']"):
                    r.click("//*[text()='Общие сведения']")
                    time.sleep(2)
                    # Пример извлечения данных – XPath следует откорректировать по реальной разметке
                    organizer = r.read("//*[contains(text(), 'Организатор:')]/following-sibling::*").strip() if r.exist("//*[contains(text(), 'Организатор:')]/following-sibling::*") else ""
                    sum_value = r.read("//*[contains(text(), 'Сумма закупки')]/following-sibling::*").strip() if r.exist("//*[contains(text(), 'Сумма закупки')]/following-sibling::*") else ""
                else:
                    organizer = ""
                    sum_value = ""
                
                new_record = {
                    "Ссылка на закупку": current_url,
                    "Лот №": file_found_lot,
                    "Организатор": organizer,
                    "Сумма закупки": sum_value
                }
                monitoring_df = monitoring_df.append(new_record, ignore_index=True)
                monitoring_df.to_excel(monitoring_file, index=False)
            else:
                print(f"Слово '{word}' не найдено ни в одном файле для объявления {unique_id}.")
            
            # Помечаем объявление как обработанное
            processed_announcements.add(unique_id)
            with open(processed_file, 'w') as f:
                for ann in processed_announcements:
                    f.write(str(ann) + "\n")
            
            # Закрываем всплывающее окно с объявлением и возвращаемся к списку объявлений
            r.keyboard("[ctrl]w")
            time.sleep(3)
            r.popup("https://goszakup.gov.kz/ru/search/announce")
            print("Переход обратно на ", back_url)
            announcement_index += 1

# После обработки всех слов и объявлений отправляем письмо с файлом мониторинга
send_email(
    subject="Мониторинг закупок",
    body="Во вложении находится файл мониторинга закупок.",
    attachment_path=monitoring_file,
    recipient_email="mentor@example.com"  # Замените на реальный адрес
)

r.close()
