from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.action_chains import ActionChains
import time
from datetime import datetime, timedelta
from selenium.common.exceptions import StaleElementReferenceException
import os
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
import subprocess
import os

# Чтение списка людей из файла
try:
    with open('/home/stas/Документы/sotrudniki.txt', 'r', encoding='utf-8') as file:
        people_list = [line.strip() for line in file if line.strip()]
        if not people_list:
            raise ValueError("Список сотрудников пуст.")
    print(f"Найдено сотрудников: {len(people_list)}")
except FileNotFoundError:
    print("Файл sotrudniki.txt не найден.")
    people_list = []
except ValueError as e:
    print(e)
    people_list = []

local_folder_path = '/home/stas/Документы/crocotime/'
remote_folder_path = 'gdrive:/1no3RzytJ_H_YPr4IbN7fCS1UQlytmss-'

# Указываем папку для скачивания файлов
download_directory = "/home/stas/Документы/crocotime"
prefs = {
    "download.default_directory": download_directory,  # Путь к папке для загрузки
    "download.prompt_for_download": False,  # Отключаем диалог "Сохранить как"
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True  # Отключаем предупреждение Chrome о возможной опасности файла
}

# Добавляем эти настройки в опции Chrome
chrome_options = Options()
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--start-maximized")  # Открыть браузер в полноэкранном режиме
#chrome_options.add_argument("--headless")  # Запускает браузер в фоновом режиме (без интерфейса)

# Инициализация драйвера с настроенными параметрами после ввода данных
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# функцимя для выгрузщки данных
def upload_files_to_google_drive(local_folder, remote_folder):
    # Команда rclone для загрузки только новых файлов
    rclone_command = [
        'rclone', 'sync', local_folder_path, remote_folder_path, '--progress'
    ]
    try:
        # Выполняем команду rclone с помощью subprocess
        result = subprocess.run(rclone_command, check=True, capture_output=True, text=True)
        print("Файлы успешно загружены на Google Диск!")
        print(result.stdout)
    except subprocess.CalledProcessError as e:
        print(f"Ошибка при загрузке файлов: {e.stderr}")

def is_weekday():
    """
    Проверяем, является ли сегодня рабочим днем (понедельник-пятница).
    Возвращает True, если сегодня будний день, и False, если выходной.
    """
    today = datetime.now().weekday()
    # Понедельник = 0, ... , Пятница = 4, Суббота = 5, Воскресенье = 6
    return today < 5  # Вернет True для будних дней

# Функция для получения текущего месяца и года в формате, как на странице (например, "Октябрь 2024")
def get_current_month_year():
    months = {
        1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель", 5: "Май", 6: "Июнь",
        7: "Июль", 8: "Август", 9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
    }
    current_date = datetime.now()
    month_name = months[current_date.month]
    year = current_date.year
    return f"{month_name}\xa0{year}"  # Используем \xa0 для пробела

# Функция для получения текущей даты в формате (например, "4 октября")
def get_current_day_month():
    months = {
        1: "января", 2: "февраля", 3: "марта", 4: "апреля", 5: "мая", 6: "июня",
        7: "июля", 8: "августа", 9: "сентября", 10: "октября", 11: "ноября", 12: "декабря"
    }
    current_date = datetime.now()
    day = current_date.day
    month_name = months[current_date.month]
    return f"{day}\xa0{month_name}"  # Используем \xa0 для пробела

# Функция для получения даты предыдущего дня или пятницы
def get_previous_day_or_friday():
    today = datetime.now()
    if today.weekday() == 0:  # Если понедельник
        previous_friday = today - timedelta(days=3)
        return previous_friday
    else:
        previous_day = today - timedelta(days=1)
        return previous_day

# Функция для расчета значения атрибута data-time для выбранной даты
def calculate_data_time(target_date):
    target_date_midnight = datetime(target_date.year, target_date.month, target_date.day)
    target_date_adjusted = target_date_midnight + timedelta(hours=3)
    return int(target_date_adjusted.timestamp() * 1000)

# Функция для поиска сотрудников и их выбора
def find_and_select_people(driver, people_list):
    try:
        wait = WebDriverWait(driver, 10)

        while True:
            rows = driver.find_elements(By.CSS_SELECTOR, "div.block-rowactivityperson")

            if not rows:
                print("Сотрудники не найдены.")
                break

            for row in rows:
                try:
                    name_element = row.find_element(By.CSS_SELECTOR, "div.block-row-captioncell-table div.block-text")
                    name_text = name_element.text.strip()
                    if name_text in people_list:
                        print(f"Найден сотрудник: {name_text}")
                        checkbox = row.find_element(By.CSS_SELECTOR, "div.block-checkboxtristate")
                        driver.execute_script("arguments[0].click();", checkbox)
                    else:
                        print(f"Сотрудник {name_text} не в списке.")
                except Exception as e:
                    print(f"Ошибка при обработке сотрудника: {str(e)}")

            driver.execute_script("window.scrollBy(0, 300);")
            time.sleep(2)

            new_rows = driver.find_elements(By.CSS_SELECTOR, "div.block-rowactivityperson")
            if len(new_rows) == len(rows):
                print("Достигнут конец списка сотрудников.")
                break

    except Exception as e:
        print(f"Ошибка при поиске сотрудников: {str(e)}")

# === Добавленные функции для переименования файлов ===

# Функция для ожидания завершения скачивания
def wait_for_downloads(download_directory):
    while True:
        if any(filename.endswith(".crdownload") for filename in os.listdir(download_directory)):
            time.sleep(1)  # Ждем пока скачивание завершится
        else:
            break

# Функция для переименования загруженного файла
def rename_downloaded_file(person_name, target_date, download_directory):
    # Ожидаем завершения скачивания всех файлов
    wait_for_downloads(download_directory)

    # Формируем имя нового файла
    date_string = target_date.strftime("%d %B")
    new_filename = f"{person_name} - {date_string}.xlsx"

    # Поиск недавно загруженного файла и его переименование
    for filename in os.listdir(download_directory):
        if filename.endswith(".xlsx"):
            old_file = os.path.join(download_directory, filename)
            new_file = os.path.join(download_directory, new_filename)
            os.rename(old_file, new_file)
            print(f"Файл переименован: {new_filename}")
            break

# Функция для обработки одного сотрудника
def process_employee(index, block):
    try:
        actions = ActionChains(driver)

        # Шаг 1: Клик по "Отработанное"
        work_summary = block.find_element(By.XPATH, ".//td[text()='Отработанное']")
        actions.move_to_element(work_summary).click().perform()
        print(f"Клик по 'Отработанное' для сотрудника {index + 1} выполнен.")

        time.sleep(1)  # Задержка для стабильности

        # Шаг 2: Клик по первой кнопке "Программы"
        first_programs_button = WebDriverWait(driver, 10).until(
            ec.element_to_be_clickable((By.XPATH, "//div[@class='block-button-value' and text()='Программы']"))
        )
        actions.move_to_element(first_programs_button).click().perform()
        print(f"Клик по первой кнопке 'Программы' для сотрудника {index + 1} выполнен.")

        time.sleep(2)  # Задержка на загрузку страницы

        # Шаг 3: Клик по второй кнопке "Программы" (раздел "Программы")
        second_programs_button = WebDriverWait(driver, 10).until(
            ec.element_to_be_clickable((By.XPATH, "//div[@class='block-button-value' and text()='Программы']"))
        )
        actions.move_to_element(second_programs_button).click().perform()
        print(f"Клик по второй кнопке 'Программы' для сотрудника {index + 1} выполнен.")

        time.sleep(2)  # Задержка для загрузки данных

        # Шаг 4: Клик по кнопке загрузки (иконка загрузки)
        download_button = WebDriverWait(driver, 10).until(
            ec.element_to_be_clickable(
                (By.XPATH, "//div[@class='block-button-icon im-icon im-icon-blue-export']"))
        )
        actions.move_to_element(download_button).click().perform()
        print(f"Клик по кнопке загрузки выполнен для сотрудника {index + 1}.")

        time.sleep(3)  # Ждем, пока откроется окно подтверждения

        # Шаг 5: Клик по кнопке "Скачать" в открывшемся окне
        confirm_download_button = WebDriverWait(driver, 10).until(
            ec.element_to_be_clickable((By.XPATH, "//div[@class='block-button-value' and text()='Скачать']"))
        )
        actions.move_to_element(confirm_download_button).click().perform()
        print(f"Клик по кнопке 'Скачать' выполнен для сотрудника {index + 1}.")

        time.sleep(5)  # Ждем завершения загрузки

        # Шаг 6: Переименование файла
        target_date = get_previous_day_or_friday()  # Дата отчета
        person_name = f"Сотрудник_{index + 1}"  # Замени на реальное имя сотрудника, если оно есть
        rename_downloaded_file(person_name, target_date, download_directory)
        print(f"Файл для сотрудника {person_name} переименован.")

        # Шаг 7: Нажимаем на кнопку "Закрыть" после завершения скачивания
        close_button = WebDriverWait(driver, 10).until(
            ec.element_to_be_clickable((By.XPATH, "//div[@class='block-button-value' and text()='Закрыть']"))
        )
        close_button.click()
        print(f"Кнопка 'Закрыть' нажата для сотрудника {person_name}.")

        time.sleep(2)  # Задержка перед обработкой следующего сотрудника

    except Exception as e:
        print(f"Ошибка при обработке сотрудника {index + 1}: {e}")


# Основной процесс обработки всех сотрудников
def process_all_employees():
    employee_blocks = driver.find_elements(By.CSS_SELECTOR, "div.block-rowactivitydaystructure-total")

    # Пройдем по всем блокам сотрудников
    for index, block in enumerate(employee_blocks):
        process_employee(index, block)

        if index != len(employee_blocks) - 1:
            # Скроллим на размер одного сотрудника (на 163 пикселя)
            driver.execute_script("window.scrollBy(0, 163);")
            #time.sleep(1)  # Задержка для стабильности скроллинга

# Проверка и закрытие вкладки "Отчеты", если она раскрыта
def close_reports_if_expanded():
    try:
        # Поиск элемента "Отчеты"
        reports_tab = driver.find_element(By.XPATH, "//div[@class='block-button-value' and text()='Отчеты']")
        reports_parent = reports_tab.find_element(By.XPATH, "..")  # Родительский элемент
        # Проверяем, раскрыта ли вкладка (проверка на наличие класса block-accordion-header_expanded)
        if "block-accordion-header_expanded" in reports_parent.get_attribute("class"):
            print("Вкладка 'Отчеты' раскрыта. Закрываем...")
            actions = ActionChains(driver)
            actions.move_to_element(reports_tab).click().perform()
            time.sleep(2)  # Задержка для закрытия вкладки
            print("Вкладка 'Отчеты' закрыта.")
        else:
            print("Вкладка 'Отчеты' уже закрыта.")
    except Exception as e:
        print(f"Ошибка при закрытии вкладки 'Отчеты': {e}")


try:
    # Открываем страницу авторизации
    login_url = "https://crocotime.infomaximum.com/?paSignIn"
    driver.get(login_url)
    print("Открыта страница:", login_url)

    # Поиск поля для ввода почты (логина)
    email_field = driver.find_element(By.NAME, "EMAIL")
    email_field.send_keys("tech@flacon-opt.ru")
    print("Email введён")

    # Поиск поля для ввода пароля
    password_field = driver.find_element(By.NAME, "PASSWORD")
    password_field.send_keys("KPyDOhW%Qo)bWs")
    print("Пароль введён")

    # Ожидание обновления состояния кнопки (чтобы она стала активной)
    time.sleep(1)  # Задержка для обработки введённых данных

    # Проверяем, активна ли кнопка
    login_button = driver.find_element(By.XPATH, "//button//span[text()='Войти']/ancestor::button")
    if login_button.is_enabled():
        # Если кнопка активна (disabled отсутствует), нажимаем её
        login_button.click()
        print("Кнопка 'Войти' нажата.")
    else:
        print("Кнопка 'Войти' заблокирована. Проверьте правильность введённых данных.")

    # Ждём несколько секунд для загрузки страницы после авторизации
    time.sleep(3)

    # Делаем скриншот страницы после выполнения действий
    screenshot_path = "screenshot_after_login.png"
    driver.save_screenshot(screenshot_path)
    print(f"Скриншот сделан и сохранён в {screenshot_path}")

    try:
        # Явное ожидание до 30 секунд, пока элемент с текстом "Отчеты" не станет кликабельным
        wait = WebDriverWait(driver, 30)

        # Проверка, открыт ли уже раздел "Отчеты"
        is_menu_open = False
        try:
            # Проверяем наличие элемента, который появляется после открытия "Отчеты"
            wait.until(ec.visibility_of_element_located((By.XPATH, "//div[text()='Детали дня']")))
            is_menu_open = True
            print("Меню 'Отчеты' уже открыто, пропускаем нажатие кнопки.")
        except:
            print("Меню 'Отчеты' не открыто, нажимаем на кнопку 'Отчеты'.")

        # Если меню не открыто, то кликаем по кнопке "Отчеты"
        if not is_menu_open:
            reports_button = wait.until(
                ec.element_to_be_clickable((By.XPATH, "//div[@class='block-button-value' and text()='Отчеты']")))

            # Прокрутка к элементу для его видимости
            driver.execute_script("arguments[0].scrollIntoView();", reports_button)

            # Использование ActionChains для более точного клика
            reports_button.click()
            print("Кнопка 'Отчеты' нажата.")

            # Ожидание появления меню после нажатия кнопки
            time.sleep(2)

        # Проверка, что меню "Отчеты" действительно раскрыто
        details_button = wait.until(ec.element_to_be_clickable((By.XPATH, "//div[text()='Детали дня']")))
        print("Меню 'Отчеты' успешно раскрыто.")

        # Получаем URL из атрибута href кнопки "Детали дня"
        details_url = details_button.find_element(By.XPATH, "..").get_attribute("href")
        print(f"URL для перехода: {details_url}")

        # Переход на страницу "Детали дня"
        driver.get(details_url)
        print("Переход на страницу 'Детали дня' выполнен.")

        # Время ожидания после перехода
        time.sleep(5)

        # Закрытие вкладки "Отчеты", если она раскрыта
        close_reports_if_expanded()

        # Проверяем, изменился ли URL после перехода
        current_url = driver.current_url
        print(f"Текущий URL: {current_url}")

        # Ожидание перехода на новую страницу
        if "day_structure" in current_url:  # Убедимся, что в URL присутствует ожидаемый сегмент
            print("Переход на страницу 'Детали дня' выполнен успешно.")
        else:
            print("Не удалось перейти на страницу 'Детали дня'. Проверьте наличие ошибок.")

        # Скриншот после перехода на новую страницу
        screenshot_path = "screenshot_after_click.png"
        driver.save_screenshot(screenshot_path)
        print(f"Скриншот сделан и сохранён: {screenshot_path}")

        # Получаем текущий месяц и год для кнопки (например, "Октябрь 2024")
        current_month_year = get_current_month_year()
        print(f"Ищем кнопку для месяца и года: {current_month_year}")

        # Ожидание элемента с динамическим месяцем и годом
        date_button = wait.until(ec.element_to_be_clickable(
            (By.XPATH, f"//div[@class='block-button-value' and text()='{current_month_year}']")))

        # Использование JavaScript для клика по кнопке
        driver.execute_script("arguments[0].click();", date_button)
        print(f"Кнопка '{current_month_year}' нажата.")

        # Ожидание изменений на странице (например, выпадение списка)
        time.sleep(2)

        # Скриншот после нажатия на кнопку
        screenshot_path = "screenshot_after_date_click.png"
        driver.save_screenshot(screenshot_path)
        print(f"Скриншот после нажатия на '{current_month_year}' сделан и сохранён: {screenshot_path}")

        # === Новый код для нажатия на кнопку "Месяц" ===

        month_button = wait.until(
            ec.element_to_be_clickable((By.XPATH, "//div[@class='block-button-value' and text()='Месяц']")))
        print("Кнопка 'Месяц' найдена и готова к клику.")

        # Использование JavaScript для клика по кнопке "Месяц"
        driver.execute_script("arguments[0].click();", month_button)
        print("Кнопка 'Месяц' нажата с использованием JavaScript.")

        # Ожидание изменений на странице после клика
        time.sleep(3)  # Увеличим задержку до 3 секунд
        print("Ожидание завершено после клика по кнопке 'Месяц'.")

        # Проверка, был ли контент страницы обновлен (можно добавить проверку появления нового элемента)
        if "Ожидаемый элемент после клика" in driver.page_source:
            print("Изменения на странице произошли после клика на кнопку 'Месяц'.")
        else:
            print("Изменений на странице не произошло после клика на кнопку 'Месяц'.")

        # Скриншот после нажатия на кнопку "Месяц"
        screenshot_path = "screenshot_after_month_click.png"
        driver.save_screenshot(screenshot_path)
        print(f"Скриншот после нажатия на кнопку 'Месяц' сделан и сохранён: {screenshot_path}")

        # === Новый код для нажатия на кнопку "День" ===

        # Поиск кнопки "День"
        day_button = wait.until(
            ec.element_to_be_clickable((By.XPATH, "//div[@class='block-button-value' and text()='День']")))

        # Использование JavaScript для клика по кнопке "День"
        driver.execute_script("arguments[0].click();", day_button)
        print("Кнопка 'День' нажата.")

        # Ожидание изменений на странице
        time.sleep(2)

        # Скриншот после нажатия на кнопку "День"
        screenshot_path = "screenshot_after_day_click.png"
        driver.save_screenshot(screenshot_path)
        print(f"Скриншот после нажатия на кнопку 'День' сделан и сохранён: {screenshot_path}")

        # Получаем текущую дату (например, "4 октября")
        current_day_month = get_current_day_month()
        print(f"Ищем кнопку для текущей даты: {current_day_month}")

        # Поиск кнопки с текущей датой (например, "4 октября")
        date_select_button = wait.until(ec.element_to_be_clickable(
            (By.XPATH, f"//div[@class='block-button-value' and contains(text(), '{current_day_month}')]")))

        # Использование JavaScript для клика по кнопке
        driver.execute_script("arguments[0].click();", date_select_button)
        print(f"Кнопка с датой '{current_day_month}' нажата.")

        # Ожидание изменений на странице (например, выпадение списка)
        time.sleep(2)

        # Скриншот после нажатия на кнопку
        screenshot_path = "screenshot_after_date_select_click.png"
        driver.save_screenshot(screenshot_path)
        print(f"Скриншот после нажатия на кнопку с датой '{current_day_month}' сделан и сохранён: {screenshot_path}")

        # Определяем, какой день выбрать: вчера или пятница
        target_date = get_previous_day_or_friday()
        print(f"Выбираем дату: {target_date.strftime('%d %B %Y')}")

        # Рассчитываем значение data-time для выбранной даты
        data_time_value = calculate_data_time(target_date)
        print(f"Ищем ячейку с data-time: {data_time_value}")

        # Поиск ячейки по атрибуту data-time и выполнение клика с использованием JavaScript
        day_cell = wait.until(ec.presence_of_element_located((By.XPATH, f"//td[@data-time='{data_time_value}']")))
        driver.execute_script("arguments[0].click();", day_cell)
        print(f"Кнопка с датой '{target_date.strftime('%d %B %Y')}' нажата.")

        screenshot_path = "screenshot_after_day_select.png"
        driver.save_screenshot(screenshot_path)
        print(f"Скриншот после выбора даты '{target_date.strftime('%d %B %Y')}' сделан и сохранён: {screenshot_path}")

        # Поиск и нажатие на кнопку "Сохранить" после выбора даты
        save_date_button = wait.until(ec.element_to_be_clickable((By.XPATH,
                                                                  "//div[@class='block-mixselectbodywithapplybutton-apply_wrap']//div[@class='block-button-value' and text()='Сохранить']")))

        # Использование ActionChains для клика по кнопке "Сохранить"
        actions = ActionChains(driver)
        actions.move_to_element(save_date_button).click().perform()
        print("Кнопка 'Сохранить' для выбора даты нажата.")

        # Ожидание, пока исчезнет всплывающее окно выбора даты, что свидетельствует об успешном сохранении
        wait.until(ec.invisibility_of_element_located(
            (By.XPATH, "//div[@class='block-mixselectbodywithapplybutton-apply_wrap']")))
        print("Выбор даты успешно сохранён.")

        # Явное ожидание элемента с текстом "Все сотрудники"
        all_employees_button = WebDriverWait(driver, 30).until(
            ec.element_to_be_clickable((By.XPATH, "//div[@class='block-button-value' and text()='Все сотрудники']"))
        )

        # Попробуем обычный клик
        all_employees_button.click()
        print("Кнопка 'Все сотрудники' нажата через обычный клик.")

        # Добавляем небольшую задержку для обработки события
        time.sleep(5)

        # Вызов функции, поиска людей
        find_and_select_people(driver, people_list)

        # === Использование ActionChains для клика по кнопке "Сохранить" ===
        for _ in range(3):  # Попробуем до 3 раз в случае ошибки StaleElementReferenceException
            try:
                # Повторный поиск кнопки "Сохранить"
                save_button = wait.until(
                    ec.element_to_be_clickable((By.XPATH, "//div[@class='block-button-value' and text()='Сохранить']")))

                actions = ActionChains(driver)
                actions.move_to_element(save_button).click().perform()
                print("Кнопка 'Сохранить' нажата через ActionChains.")

                # Добавляем небольшую задержку для обработки события
                time.sleep(2)

                # Ожидание, пока блок с сотрудниками исчезнет, что подтверждает успешное выполнение
                wait.until(ec.invisibility_of_element_located(
                    (By.XPATH, "//div[@class='block-popupchoosepersonswithfilter block-popupform block-popup']")))
                print("Блок 'Сотрудники' исчез, действие завершено.")

                # Скриншот после нажатия на кнопку "Сохранить"
                screenshot_path = "screenshot_after_save_button_click.png"
                driver.save_screenshot(screenshot_path)
                print(f"Скриншот сделан и сохранён: {screenshot_path}")
                break  # Успешное выполнение, выход из цикла
            except StaleElementReferenceException:
                print("Элемент устарел, пытаемся снова...")

        process_all_employees()



    except Exception as e:
        print(f"Произошла ошибка: {str(e)}")
finally:
    time.sleep(5)
    driver.quit()
    print("Браузер закрыт.")

# Проверяем, если сегодня будний день
if is_weekday():
    # Запуск загрузки файлов
    upload_files_to_google_drive(local_folder_path, remote_folder_path)
else:
    print("Сегодня выходной, загрузка файлов не выполняется.")
