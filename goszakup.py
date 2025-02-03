import os
import pandas as pd
import rpa as rpa
import pdfplumber
import codecs
from docx import Document

def search_word_in_file(file_path, search_word):
    """Функция для поиска слова в файле (PDF или DOCX)"""
    file_extension = file_path.split('.')[-1].lower()

    try:
        if file_extension == "pdf":
            print(f"Открываем PDF-файл: {file_path}")
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text and search_word.lower() in text.lower():
                        print(f"✅ Слово '{search_word}' найдено в файле {file_path}")
                        return True
            print(f"❌ Слово '{search_word}' не найдено в {file_path}")

        elif file_extension in ["doc", "docx"]:
            print(f"Открываем Word-файл: {file_path}")
            doc = Document(file_path)
            for paragraph in doc.paragraphs:
                if search_word.lower() in paragraph.text.lower():
                    print(f"✅ Слово '{search_word}' найдено в файле {file_path}")
                    return True
            print(f"❌ Слово '{search_word}' не найдено в {file_path}")

        else:
            print(f"⚠️ Неподдерживаемый формат файла: {file_path}")

    except Exception as e:
        print(f"❌ Ошибка обработки файла {file_path}: {e}")

    return False

def process_tags_file():
    # Проверяем наличие файла
    file_name = 'Теги.xlsx'
    if not os.path.exists(file_name):
        print(f"Файл {file_name} не найден в текущей директории.")
        return

    # Открываем Excel файл
    try:
        data = pd.read_excel(file_name)  # Загружаем файл с заголовками колонок
    except Exception as e:
        print(f"Ошибка при открытии файла: {e}")
        return

    # Проверяем наличие данных в первой колонке (колонка A)
    if data.empty or data.shape[1] < 1:
        print("Файл пустой или колонка 'A' отсутствует.")
        return

    # Получаем список слов из первой колонки (A), исключая заголовок
    words = data.iloc[0:, 0].dropna().tolist()

    if not words:
        print("Колонка 'A' пуста. Завершаем процесс.")
        return

    print("Слова для обработки:")
    print(words)

    # Настраиваем RPA
    print("Инициализация RPA...")
    rpa.init(visual_automation=True)

    # Переходим на сайт Госзакупок
    print("Открываем сайт Госзакупок...")
    rpa.url("https://goszakup.gov.kz/ru/search/announce")

    for word in words:
        try:
            print(f"Обрабатываем слово: {word}")

            # Заполняем поле "Наименование объявления"
            print("Вводим наименование объявления...")
            rpa.type("//*[@id='in_name']", word)

            # Выбираем "Способ закупки"
            print("Выбираем способ закупки...")
            rpa.click("//*[@data-id='s2_method']")
            rpa.click("//*[text()='Запрос ценовых предложений']")
            rpa.click("//*[text()='Открытый конкурс']")

            # Устанавливаем статус "Завершено"
            print("Выбираем статус 'Завершено'...")
            rpa.click("//*[@data-id='s2_status']")
            rpa.click("//*[text()='Завершено']")

            # Устанавливаем сумму закупки "С"
            print("Устанавливаем сумму закупки...")
            rpa.type("//*[@id='in_amount_from']", "1000000")

            # Устанавливаем "Финансовый год"
            print("Выбираем финансовый год...")
            rpa.click("//*[@data-id='s2_year']")
            rpa.click("//*[text()='2021']")
            rpa.click("//*[text()='2020']")

            # Нажимаем кнопку "Найти"
            print("Нажимаем кнопку 'Найти'...")
            rpa.click("//button[@name='smb' and contains(@class, 'btn-success')]")
            rpa.wait(5)

            # Открываем объявление
            first_announcement_xpath = "//*[@id='search-result']//tr[1]//a"
            if rpa.exist(first_announcement_xpath):
                print("Объявление найдено. Открываем...")
                rpa.click(first_announcement_xpath)
                rpa.wait(5)

                announcement_url = rpa.read(first_announcement_xpath + "/@href")
                full_url = f"https://goszakup.gov.kz{announcement_url}" if announcement_url.startswith("/") else announcement_url
                rpa.popup(full_url)

                # Переход во вкладку "Документация"
                documentation_xpath = "//*[text()='Документация']"
                if rpa.exist(documentation_xpath):
                    rpa.click(documentation_xpath)
                    rpa.wait(2)

                    # Нажатие на кнопку "Перейти"
                    tech_spec_xpath = "//*[contains(text(), 'Техническая спецификация')]/following-sibling::td/button[contains(text(), 'Перейти')]"
                    if rpa.exist(tech_spec_xpath):
                        rpa.click(tech_spec_xpath)
                        rpa.wait(3)
                        # Обрабатываем файлы
                        file_links = rpa.read("//*[@id='ModalShowFilesBody']//a[contains(@href, 'download_file')]/@href")
                        if isinstance(file_links, str):
                            file_links = [file_links]
                        
                        for file_link in file_links:
                            file_url = f"https://v3bl.goszakup.gov.kz{file_link}" if file_link.startswith("/") else file_link
                            print(f"Скачиваем файл: {file_url}")
                            rpa.url(file_url)
                            rpa.wait(3)
                            
                            file_name = rpa.read(f"//*[@id='ModalShowFilesBody']//a[@href='{file_link}']")
                            # Декодируем юникод-строку
                            file_name = codecs.decode(file_name, 'unicode_escape')
                            file_path = os.path.join(os.getcwd(), file_name)
                            
                            print(file_name, file_path)
                            if search_word_in_file(file_path, word):
                                print("✅ Слово найдено, переходим к следующему этапу...")
                                break
                            else:
                                print("❌ Слово не найдено, продолжаем проверку других файлов...")
        except Exception as e:
            print(f"Ошибка при обработке слова '{word}': {e}")

    # Завершаем работу RPA
    print("Завершаем работу RPA...")
    rpa.close()

if __name__ == "__main__":
    process_tags_file()
