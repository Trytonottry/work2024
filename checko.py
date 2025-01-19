import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from threading import Thread

def extract_revenue(inn, excel_file, inn_column, revenue_column, df):
  """
  Извлекает выручку с сайта checko.ru по заданному ИНН.
  """
  try:
    # Инициализация драйвера браузера
    driver = webdriver.Chrome()

    # Открытие сайта checko.ru
    driver.get("https://checko.ru/")

    # Ожидание загрузки страницы
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "search-form")))

    # Ввод ИНН в поисковую строку
    search_field = driver.find_element(By.ID, "search-form")
    search_field.send_keys(inn)

    # Нажатие на кнопку поиска
    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

    # Ожидание загрузки страницы с результатами
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "company-card")))

    # Получение значения выручки
    revenue_element = driver.find_element(By.XPATH, "//dl[@class='company-card__info']//dd[1]")
    revenue = revenue_element.text

    # Закрытие браузера
    driver.quit()

    # Запись выручки в DataFrame
    row_index = df[df[inn_column] == inn].index[0]
    df.loc[row_index, revenue_column] = revenue

  except TimeoutException:
    print(f"Ошибка: Страница не загрузилась вовремя для ИНН: {inn}")

  except Exception as e:
    print(f"Ошибка: {e}")

def main():
  # Задание параметров
  excel_file = "data.xlsx"
  inn_column = "ИНН"
  revenue_column = "Выручка"

  # Чтение данных из Excel файла
  df = pd.read_excel(excel_file)

  # Создание потоков для каждого ИНН
  threads = []
  for i in range(len(df)):
    inn = df.loc[i, inn_column]
    thread = Thread(target=extract_revenue, args=(inn, excel_file, inn_column, revenue_column, df))
    threads.append(thread)
    thread.start()

  # Ожидание завершения всех потоков
  for thread in threads:
    thread.join()

  # Сохранение обновленных данных в Excel файл
  df.to_excel(excel_file, index=False)

if __name__ == "__main__":
  main()
