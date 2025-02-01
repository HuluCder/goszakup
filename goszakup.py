import os
import pandas as pd
import rpa as rpa

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
    words = data.iloc[1:, 0].dropna().tolist()

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

            # Ждём несколько секунд для загрузки результатов
            print("Ждём загрузки результатов поиска...")
            rpa.wait(5)

            # Открываем первое объявление из списка результатов
            print("Проверяем наличие объявления...")
            first_announcement_xpath = "//*[@id='search-result']//tr[1]//a"
            if rpa.exist(first_announcement_xpath):
                print("Объявление найдено. Открываем...")
                rpa.click(first_announcement_xpath)
                rpa.wait(5)  # Ждём загрузки новой вкладки

                # Кликаем в область страницы для установки фокуса
                print("Кликаем по странице для установки фокуса...")
                rpa.click("/html/body")
                (rpa.keyboard('[pagedown]'))
                rpa.wait(2)

                # Проверяем наличие вкладки 'Документация'
                print("Проверяем наличие вкладки 'Документация'...")
                documentation_xpath = "//*[text()='Документация']"
                if rpa.exist(documentation_xpath):
                    print("Переключение на новую вкладку успешно.")

                    # Переходим во вкладку "Документация"
                    print("Переходим во вкладку 'Документация'...")
                    rpa.click(documentation_xpath)
                    rpa.wait(2)

                    # Ищем "Техническая спецификация"
                    print("Ищем 'Техническая спецификация'...")
                    tech_spec_xpath = "//*[contains(text(), 'Техническая спецификация')]/following-sibling::td/button[contains(text(), 'Перейти')]"
                    if rpa.exist(tech_spec_xpath):
                        print(f"Техническая спецификация найдена. Открываем...")
                        rpa.click(tech_spec_xpath)
                        rpa.wait(3)
                    else:
                        print(f"Техническая спецификация отсутствует. Пропускаем объявление.")
                else:
                    print("Ошибка: Робот не переключился на новую вкладку.")

            else:
                print(f"Объявление для слова '{word}' не найдено.")

        except Exception as e:
            print(f"Ошибка при обработке слова '{word}': {e}")

    # Завершаем работу RPA
    print("Завершаем работу RPA...")
    rpa.close()

# Запуск функции
if __name__ == "__main__":
    process_tags_file()
