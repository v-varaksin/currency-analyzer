import tkinter as tk  # для создания окон и кнопок
from tkinter import messagebox  # для всплывающих окошек с сообщениями
from tkinter import filedialog  # для выбора файлов на компьютере
from tkinter import simpledialog  # для окошек с полем ввода текста
import pandas as pd  # для работы с таблицами (как Excel)
import numpy as np  # для быстрых математических вычислений
import matplotlib.pyplot as plt  # для рисования графиков
from datetime import datetime  # для работы с датами (переводит строку в дату)
from datetime import timedelta  # для вычитания дней из даты (5 лет назад)
import requests  # для отправки запросов в интернет
import xml.etree.ElementTree as ET  # для разбора XML (язык сайта ЦБ РФ)


class CurrencyAnalyzer:
    """Программа для анализа курса иены"""

    def __init__(self):
        """Создаем окно и кнопки"""

        # Переменная для хранения таблицы с данными о курсах
        self.data_frame = None

        # Создаем главное окно
        self.main_window = tk.Tk()
        self.main_window.title("Курс японской иены (вариант 5)")
        self.main_window.geometry("450x550")

        # Создаем заголовок окна
        window_title = tk.Label(self.main_window, text="Анализ курса японской иены",
                                font=("Arial", 14, "bold"))
        window_title.pack(pady=10)

        # КНОПКА 1: Загрузить данные из интернета
        button_load = tk.Button(self.main_window, text="1. ЗАГРУЗИТЬ ДАННЫЕ",
                                command=self.load_data, bg="lightblue", height=2)
        button_load.pack(fill="x", padx=30, pady=5)

        # КНОПКА 2: Проанализировать данные за 5 лет
        button_analyze = tk.Button(self.main_window, text="2. АНАЛИЗ ЗА 5 ЛЕТ",
                                   command=self.analyze, bg="lightgreen", height=2)
        button_analyze.pack(fill="x", padx=30, pady=5)

        # КНОПКА 3: Показать курс на конкретную дату
        button_get_rate = tk.Button(self.main_window, text="3. КУРС НА ДАТУ",
                                    command=self.get_rate, bg="yellow", height=2)
        button_get_rate.pack(fill="x", padx=30, pady=5)

        # КНОПКА 4: Построить график за выбранный период
        button_period = tk.Button(self.main_window, text="4. ГРАФИК ЗА ПЕРИОД",
                                  command=self.plot_period, bg="orange", height=2)
        button_period.pack(fill="x", padx=30, pady=5)

        # КНОПКА 5: Сохранить все данные в EXCEL/Яндекс Диск
        button_save = tk.Button(self.main_window, text="5. СОХРАНИТЬ В EXCEL/ЯНДЕКС ДИСК",
                                command=self.save_excel, bg="gray", fg="white", height=2)
        button_save.pack(fill="x", padx=30, pady=5)

        # КНОПКА 6: Прогноз курса на 30 дней
        button_forecast = tk.Button(self.main_window, text="6. ПРОГНОЗ НА 1 ГОД",
                                    command=self.forecast, bg="lightpink", height=2)
        button_forecast.pack(fill="x", padx=30, pady=5)

        # КНОПКА ВЫХОДА: Закрыть программу
        button_exit = tk.Button(self.main_window, text="ВЫХОД",
                                command=self.main_window.quit, bg="red", fg="white", height=2)
        button_exit.pack(fill="x", padx=30, pady=5)

        # Запускаем главный цикл программы (ожидание действий пользователя)
        self.main_window.mainloop()

    # 1. ФУНКЦИЯ ЗАГРУЗКИ ДАННЫХ
    def load_data(self):
        """Загружает курс японской иены с сайта Центрального Банка России"""

        # Определяем конечную дату (сегодняшний день)
        end_date = datetime.now()

        # Определяем начальную дату
        start_date = end_date - timedelta(days=5 * 365)

        # Создаем пустые списки для хранения дат и курсов
        dates_list = []
        rates_list = []

        # Начинаем с начальной даты
        current_date = start_date

        # Цикл: пока не дошли до конечной даты
        while current_date <= end_date:

            # Определяем последний день текущего месяца
            if current_date.month == 12:
                next_month_date = datetime(current_date.year + 1, 1, 1)
            else:
                next_month_date = datetime(current_date.year, current_date.month + 1, 1)

            # Форматируем даты для запроса к API ЦБ РФ
            from_date_string = "/".join(str(current_date.date()).split("-")[::-1])
            last_day_of_month = next_month_date - timedelta(days=1)
            to_date_string = "/".join(str(last_day_of_month.date()).split("-")[::-1])

            # Формируем URL адрес для запроса (R01235 = код японской иены)
            request_url = (f"http://www.cbr.ru/scripts/XML_dynamic.asp?date_req1="
                           f"{from_date_string}&date_req2={to_date_string}&VAL_NM_RQ=R01235")

            response = requests.get(request_url)
            xml_root = ET.fromstring(response.text)

            for record in xml_root.findall('Record'):
                dates_list.append(record.get('Date'))
                rates_list.append(float(record.find('Value').text.replace(',', '.')))

            # Переходим к следующему месяцу
            current_date = next_month_date

        # Создаем таблицу данных с помощью Pandas
        self.data_frame = pd.DataFrame({'date': dates_list, 'rate': rates_list})

        # Превращаем строки в настоящие даты
        self.data_frame['date'] = pd.to_datetime(self.data_frame['date'], format="%d.%m.%Y")

        # Сортируем по объектам datetime
        self.data_frame = self.data_frame.sort_values('date').reset_index(drop=True)

        # Добавляем колонку с годом
        self.data_frame['year'] = self.data_frame['date'].dt.year

    #  2. ФУНКЦИЯ АНАЛИЗА ЗА 5 ЛЕТ
    def analyze(self):
        """Анализирует курс: вычисляет максимум, минимум и среднее за каждый год"""
        # Проверка: если данные еще не загружены
        if self.data_frame is None:
            messagebox.showwarning("Ошибка", "Сначала загрузите данные (кнопка 1)")
            return
        # Получаем все уникальные годы из таблицы
        all_years = sorted(self.data_frame['year'].unique())
        # Берем только последние 5 лет
        last_five_years = all_years[-5:]
        # Список для хранения результатов анализа
        analysis_results = []

        # Анализируем каждый год отдельно
        for current_year in last_five_years:
            # Выбираем данные только за текущий год
            year_data = self.data_frame[self.data_frame['year'] == current_year]

            # Превращаем колонку с курсами в массив NumPy
            rates_array = year_data['rate'].values

            # Вычисляем статистику с помощью NumPy
            maximum_rate = np.max(rates_array)
            minimum_rate = np.min(rates_array)
            mean_rate = np.mean(rates_array)

            # Добавляем результат в список
            analysis_results.append([current_year, maximum_rate, minimum_rate, mean_rate])
        # Сохраняем результаты для использования в других функциях (например, для Excel)
        self.saved_analysis_results = analysis_results

        # СТРОИМ ГРАФИК 1: Изменение курса за 5 лет
        five_years_data = self.data_frame[self.data_frame['year'].isin(last_five_years)]
        # Сортируем данные по дате
        five_years_data = five_years_data.sort_values('date')
        fig, ax = plt.subplots(figsize=(12, 6))
        # Рисуем линию
        ax.plot(five_years_data['date'], five_years_data['rate'], 'b-', linewidth=1, label='Курс')

        # Находим максимальное и минимальное значение за весь период с помощью NumPy
        all_rates_array = five_years_data['rate'].values
        global_maximum = np.max(all_rates_array)
        global_minimum = np.min(all_rates_array)
        # Находим строки с максимальным и минимальным значением
        max_row = five_years_data[five_years_data['rate'] == global_maximum].iloc[0]
        min_row = five_years_data[five_years_data['rate'] == global_minimum].iloc[0]

        # Отмечаем точку максимума красным кружком
        plt.plot(max_row['date'], global_maximum, 'ro', markersize=10, label=f'Макс: {global_maximum:.2f}')
        # Отмечаем точку минимума зеленым кружком
        plt.plot(min_row['date'], global_minimum, 'go', markersize=10, label=f'Мин: {global_minimum:.2f}')

        # Настройки графика
        plt.title('Курс японской иены за 5 лет')
        plt.xlabel('Дата')
        plt.ylabel('Курс (руб)')
        plt.legend()
        plt.grid(True, alpha=0.3)
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()

        # СТРОИМ ДИАГРАММУ 2: Средние значения по годам

        # Извлекаем годы и средние значения из результатов
        years_list = [row[0] for row in analysis_results]
        means_list = [row[3] for row in analysis_results]

        # Создаем новое окно для диаграммы
        plt.figure(figsize=(8, 5))

        # Рисуем столбчатую диаграмму
        bars = plt.bar(years_list, means_list, color='skyblue')

        # Подписываем значения над каждым столбцом
        plt.bar_label(bars, fmt='%.1f', padding=5)

        # Настройки диаграммы
        plt.title('Средний курс по годам')
        plt.xlabel('Год')
        plt.ylabel('Средний курс (руб)')
        plt.grid(True, alpha=0.3, axis='y')
        plt.tight_layout()
        plt.show()

    # 3. ФУНКЦИЯ КУРС НА ДАТУ
    def get_rate(self):
        """Показывает курс японской иены на выбранную пользователем дату"""
        # Проверка: если данные еще не загружены
        if self.data_frame is None:
            messagebox.showwarning("Ошибка", "Сначала загрузите данные")
            return

        # Спрашиваем у пользователя дату
        user_date_string = simpledialog.askstring("Дата", "Введите дату (ГГГГ-ММ-ДД):")
        # Если пользователь ввел дату (не нажал Отмена)
        if user_date_string is not None:

            # Превращаем строку в объект даты
            target_date = datetime.strptime(user_date_string, "%Y-%m-%d")
            # Ищем точное совпадение в таблице
            exact_match = self.data_frame[self.data_frame['date'].dt.date == target_date.date()]
            # Если нашли точную дату
            if len(exact_match) > 0:
                rate_value = exact_match.iloc[0]['rate']
                messagebox.showinfo("Результат", f"Курс на {user_date_string}: {rate_value:.2f} руб.")

            # Если точной даты нет
            else:
                # Создаем временную колонку с разницей в днях
                self.data_frame['days_difference'] = abs((self.data_frame['date'] - target_date).dt.days)

                # Находим строку с минимальной разницей
                closest_row = self.data_frame.loc[self.data_frame['days_difference'].idxmin()]

                # Удаляем временную колонку
                self.data_frame = self.data_frame.drop(columns=['days_difference'])

                # Формируем сообщение
                closest_date_string = closest_row['date'].strftime('%Y-%m-%d')
                closest_rate_string = f"{closest_row['rate']:.2f}"
                messagebox.showinfo("Результат",
                                    f"Нет данных на {user_date_string}\nБлижайшая дата:"
                                    f" {closest_date_string}, курс: {closest_rate_string}")

    #  4. ФУНКЦИЯ ГРАФИК ЗА ПЕРИОД
    def plot_period(self):
        """Строит график курса за выбранный пользователем период времени"""

        # Проверка: если данные еще не загружены
        if self.data_frame is None:
            messagebox.showwarning("Ошибка", "Сначала загрузите данные")
            return

        # Спрашиваем начальную дату периода
        start_date_string = simpledialog.askstring("Период", "Начальная дата (ГГГГ-ММ-ДД):")

        # Если пользователь нажал Отмена
        if start_date_string is None:
            return

        # Спрашиваем конечную дату периода
        end_date_string = simpledialog.askstring("Период", "Конечная дата (ГГГГ-ММ-ДД):")

        # Если пользователь нажал Отмена
        if end_date_string is None:
            return

        # Превращаем строки в объекты дат
        start_date = datetime.strptime(start_date_string, "%Y-%m-%d")
        end_date = datetime.strptime(end_date_string, "%Y-%m-%d")

        # Выбираем данные за указанный период
        period_data = self.data_frame[
            (self.data_frame['date'] >= start_date) & (self.data_frame['date'] <= end_date)]
        period_data = period_data.sort_values('date')

        # Проверка: есть ли данные за этот период
        if len(period_data) == 0:
            messagebox.showwarning("Нет данных", "За этот период нет данных")
            return

        # Создаем новое окно для графика
        plt.figure(figsize=(10, 5))

        # Рисуем зеленую линию с точками
        plt.plot(period_data['date'], period_data['rate'], 'g-', linewidth=1.5)

        # Вычисляем статистику с помощью NumPy
        rates_array = period_data['rate'].values
        minimum_rate = np.min(rates_array)
        maximum_rate = np.max(rates_array)
        mean_rate = np.mean(rates_array)

        # Настройки графика
        plt.title(f'Курс иены с {start_date_string} по {end_date_string}')
        plt.xlabel('Дата')
        plt.ylabel('Курс (руб)')
        plt.grid(True, alpha=0.3)
        plt.xticks(rotation=45)

        # Добавляем текстовую подпись со статистикой
        statistics_text = f"Мин: {minimum_rate:.2f}  Макс: {maximum_rate:.2f}  Сред: {mean_rate:.2f}"
        plt.text(0.02, 0.98, statistics_text, transform=plt.gca().transAxes,
                 fontsize=9, bbox=dict(boxstyle='round', facecolor='wheat'))

        plt.tight_layout()
        plt.show()

    # 5. ФУНКЦИЯ СОХРАНЕНИЯ (ЛОКАЛЬНО / ОБЛАКО)
    def save_excel(self):
        """Сохраняет данные локально или в Яндекс.Диск"""
        if self.data_frame is None:
            messagebox.showwarning("Ошибка", "Нет данных для сохранения")
            return

        # Выбор места сохранения
        save_to_cloud = messagebox.askyesno("Выбор хранилища",
                                            "Сохранить в облако (Яндекс.Диск)?"
                                            "\nНет = локальный диск")

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        # Сохраняем исходные данные
        self.data_frame.to_excel(file_path, sheet_name='Данные', index=False)

        # Сохраняем анализ, если он был
        if hasattr(self, 'saved_analysis_results'):
            results_df = pd.DataFrame(self.saved_analysis_results,
                                      columns=['Год', 'Максимум', 'Минимум', 'Средний'])
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
                results_df.to_excel(writer, sheet_name='Анализ', index=False)

        # Если выбрали облако, загружаем файл на Яндекс.Диск
        if save_to_cloud:
            self.upload_to_yandex_disk(file_path)

    def upload_to_yandex_disk(self, file_path):
        """Загружает файл в Яндекс.Диск через API"""
        TOKEN = "y0__wgBEMz1ia4FGLK7QSCgityuF_IWhbe_eZWW5_r8CaZj3aMWqR2Q"

        filename = "Forecast.xlsx"
        upload_url = (f"https://cloud-api.yandex.net/v1/"
                      f"disk/resources/upload?path=app:/{filename}&overwrite=true")

        headers = {"Authorization": f"OAuth {TOKEN}"}
        response = requests.get(upload_url, headers=headers)
        upload_link = response.json().get("href")

        with open(file_path, "rb") as f:
            requests.put(upload_link, data=f, headers=headers)

        messagebox.showinfo("Облако", "Файл успешно загружен на Яндекс.Диск!")

    # 6. ФУНКЦИЯ ПРОГНОЗА НА 1 ГОД
    def forecast(self):
        """Строит прогноз курса японской иены на 1 год вперед"""

        # Проверка: если данные еще не загружены
        if self.data_frame is None:
            messagebox.showwarning("Ошибка", "Сначала загрузите данные")
            return

        # Берем данные за последний год и сортируем по дате
        recent = self.data_frame.tail(365).sort_values('date')

        # Вычисление прогноза
        x = np.arange(len(recent))
        y = recent['rate'].values

        # Строим линейную регрессию
        coeffs = np.polyfit(x, y, deg=1)
        poly = np.poly1d(coeffs)

        # Прогноз на 365 дней
        future_x = np.arange(len(recent), len(recent) + 365)
        forecast_values = poly(future_x)

        # Создаем даты для будущего периода
        last_date = recent['date'].max()
        future_dates = [last_date + timedelta(days=i + 1) for i in range(365)]

        # Построение графика
        plt.figure(figsize=(10, 5))

        # Рисуем историю (синяя линия)
        plt.plot(recent['date'], recent['rate'], 'b-', label='История', linewidth=1)

        # Рисуем прогноз (красный пунктир)
        plt.plot(future_dates, forecast_values, 'r--', label='Прогноз (1 год)', linewidth=2)

        # Настройки графика
        plt.title('Прогноз курса японской иены на 1 год')
        plt.xlabel('Дата')
        plt.ylabel('Курс (руб)')
        plt.legend()
        plt.grid(True, alpha=0.3)
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()


#  ЗАПУСК ПРОГРАММЫ
CurrencyAnalyzer()
