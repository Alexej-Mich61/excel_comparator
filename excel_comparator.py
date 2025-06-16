import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap import Style
from tkinter import filedialog, messagebox
import os
import re

# Применяем тему "solar"
style = Style(theme='solar')
root = style.master
root.title("Excel Comparator")
root.geometry("800x600")

def load_excel_file():
    # Загрузка пути к файлу Excel
    file_path = filedialog.askopenfilename(filetypes=[("Файлы Excel", "*.xlsx *.xls")])
    if file_path:
        try:
            # Определяем движок в зависимости от расширения файла
            engine = 'openpyxl' if file_path.lower().endswith('.xlsx') else 'xlrd'
            df = pd.read_excel(file_path, engine=engine)
            # Если используется xlrd и есть проблемы с кодировкой, пытаемся декодировать
            if engine == 'xlrd':
                for column in df.columns:
                    if df[column].dtype == 'object':  # Проверяем текстовые столбцы
                        df[column] = df[column].apply(
                            lambda x: str(x).encode('iso-8859-1').decode('cp1251', errors='ignore') if pd.notna(
                                x) else x)
            return df, file_path
        except Exception as e:
            messagebox.showerror("Ошибочка", f"Не удалось загрузить файл: {e}\nПроверьте кодировку или формат файла.")
            return None, file_path
    return None, None


def display_dataframe(df, tree, file_path, row_counter_label, is_first_tree=False):
    # Очистка дерева от предыдущих данных
    tree.delete(*tree.get_children())
    if df is None:
        tree.insert("", "end", values=("Данные не загружены",))
        if row_counter_label:
            row_counter_label.config(text="Строк: 0")
        return
    columns = list(df.columns)
    tree["columns"] = columns
    tree["show"] = "headings"
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, anchor="w")

    # Сортировка данных
    if is_first_tree:
        # Для первого Treeview (file1): сортировка по первому числу в столбце A
        def get_first_number(value):
            if pd.isna(value):
                return float('inf')  # Бесконечность для NaN
            numbers = re.split(r'[, ]+', str(value).strip())
            first_num = next((int(num) for num in numbers if re.match(r'^\d+$', num)), float('inf'))
            return first_num

        df_sorted = df.sort_values(by=df.columns[0], key=lambda x: x.apply(get_first_number))
    else:
        # Для второго Treeview (file2): сортировка по первому столбцу как есть
        df_sorted = df.sort_values(by=df.columns[0])

    # Вставка отсортированных данных
    for index, row in df_sorted.iterrows():
        values = ["" if pd.isna(row[col]) else row[col] for col in columns]
        tree.insert("", "end", values=values)
    tree.insert("", "end", values=[f"Файл: {os.path.basename(file_path)}"])
    if row_counter_label:
        row_counter_label.config(text=f"Строк: {len(df)}")

def export_to_excel(tree, file1_path, file2_path=None):
    # Извлечение данных из Treeview
    columns = tree["columns"]
    if not columns or any(not c for c in columns):  # Проверка на пустые или некорректные столбцы
        messagebox.showwarning("Э! Алё!", "Сначала загрузите данные для экспорта!")
        return

    data = []
    for item in tree.get_children():
        values = tree.item(item, "values")
        # Исключаем последнюю строку с именем файла
        if not values or not any(v.startswith("Файл:") for v in values):
            data.append(values)

    if not data:
        messagebox.showwarning("Э! Алё!", "Нет данных для экспорта!")
        return

    # Создание DataFrame
    df = pd.DataFrame(data, columns=columns)

    # Сохранение в файл
    default_name = "exported_data.xlsx"
    if file1_path and not file2_path:
        default_name = os.path.splitext(os.path.basename(file1_path))[0] + "_exported.xlsx"
    elif file2_path:
        default_name = os.path.splitext(os.path.basename(file2_path))[0] + "_exported.xlsx"

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")],
                                             initialfile=default_name)
    if save_path:
        try:
            df.to_excel(save_path, index=False, engine='openpyxl')
            messagebox.showinfo("Ништяк", f"Данные экспортированы в {save_path}")
        except Exception as e:
            messagebox.showerror("Капут", f"Не удалось экспортировать данные: {e}")

def apply_filter_report(df):
    if df is None:
        return None

    # Функция для нормализации значений в столбце A
    def normalize_column_a(value):
        if pd.isna(value):
            return ""
        # Разделение по запятым и пробелам
        value_str = re.split(r'[, ]+', str(value).strip())
        # Фильтрация нечисловых частей и сортировка чисел по возрастанию
        numbers = sorted([num for num in value_str if re.match(r'^\d+$', num)])
        if numbers:
            return ", ".join(numbers)  # Форматирование как "num1, num2" в порядке возрастания
        return value

    # Применение нормализации к столбцу A (индекс 0)
    df.iloc[:, 0] = df.iloc[:, 0].apply(normalize_column_a)

    # Добавление столбца для отметки дубликатов
    df['Дубликат'] = ''  # Изначально пустой столбец

    # Проверка на дублирующиеся значения в столбце A только для числовых значений
    numeric_mask = df.iloc[:, 0].str.contains(r'\d', na=False)
    if numeric_mask.any():
        numeric_values = df.loc[numeric_mask, df.columns[0]]
        duplicates_mask = numeric_values.duplicated(keep=False)
        if duplicates_mask.any():
            # Отмечаем все строки с дубликатами меткой "ДУБЛЬ"
            df.loc[numeric_mask & duplicates_mask, 'Дубликат'] = 'ДУБЛЬ'

    # Очистка столбцов H (индекс 7) и I (индекс 8), оставляя только числовые значения
    for col_idx in [7, 8]:  # Индексы столбцов H и I
        if col_idx < len(df.columns):  # Проверка, чтобы избежать IndexError
            df.iloc[:, col_idx] = pd.to_numeric(df.iloc[:, col_idx], errors='coerce')
            df.iloc[:, col_idx] = df.iloc[:, col_idx].where(df.iloc[:, col_idx].notna(), None)

    # Фильтрация: столбец A содержит хотя бы одно числовое значение,
    # и столбцы H или I содержат числовые значения после очистки
    df = df[
        (df.iloc[:, 0].str.contains(r'\d', na=False)) &  # Проверка наличия цифр в столбце A
        ((df.iloc[:, 7].notna()) | (df.iloc[:, 8].notna()))  # Проверка наличия чисел в H или I
    ]
    return df

def apply_filter_pcn_data(df):
    if df is None:
        return None

    # Получение количества столбцов в DataFrame
    max_index = len(df.columns) - 1

    # Определение желаемых индексов столбцов (A=0, G=6, O=14, AL=37)
    desired_indices = [0, 6, 14, 37]

    # Фильтрация индексов, выходящих за пределы
    columns_to_keep = [df.columns[i] for i in desired_indices if i <= max_index]

    # Если нет валидных столбцов, оставить хотя бы столбец A (индекс 0)
    if not columns_to_keep:
        columns_to_keep = [df.columns[0]]

    df = df[columns_to_keep]

    # Фильтрация: оставить только строки, где столбец A содержит числовое значение
    df = df[pd.to_numeric(df.iloc[:, 0], errors='coerce').notna()]

    return df

def filter_in_progress_no_contract(file2, tree2, row_counter_label2, file1, file1_path, original_file2):
    # Функция для фильтрации "В работе без договора" во втором Treeview
    if file2 is None:
        messagebox.showwarning("Э! Алё!", "Сначала загрузите данные ПЦН!")
        return

    # Сброс фильтрации по "Отключен с договором"
    file2 = original_file2.copy() if original_file2 is not None else file2

    # Условие 1: второй столбец (индекс 1) должен быть равен "Работа"
    condition1 = file2.iloc[:, 1] == "Работа"

    # Условие 2: числовое значение из первого столбца не должно встречаться в первом столбце первого Treeview
    if file1 is not None:
        # Извлечение всех чисел из первого столбца первого Treeview
        def extract_numbers(value):
            if pd.isna(value):
                return []
            value_str = str(value).replace(" ", "").split(",")
            return [num for num in value_str if re.match(r'^\d+$', str(num).strip())]

        report_numbers = set()
        for value in file1.iloc[:, 0]:
            report_numbers.update(extract_numbers(value))

        # Проверка, что значение первого столбца не содержится в report_numbers
        def check_exclusion(value):
            if pd.isna(value):
                return False
            numbers = extract_numbers(value)
            return not any(num in report_numbers for num in numbers)

        condition2 = file2.iloc[:, 0].apply(check_exclusion)
    else:
        condition2 = pd.Series([True] * len(file2))  # Если первый Treeview пуст, пропускаем условие 2

    # Применение фильтров
    file2 = file2[condition1 & condition2]
    display_dataframe(file2, tree2, file1_path if file1_path else "Файл не загружен", row_counter_label2)

def filter_deactivated_with_contract(file2, tree2, row_counter_label2, file1, file1_path, original_file2):
    # Функция для фильтрации "Отключен с договором" во втором Treeview
    if file2 is None:
        messagebox.showwarning("Э! Алё!", "Сначала загрузите данные ПЦН!")
        return

    # Сброс фильтрации по "В работе без договора"
    file2 = original_file2.copy() if original_file2 is not None else file2

    # Условие 1: второй столбец (индекс 1) не должен быть равен "Работа"
    condition1 = file2.iloc[:, 1] != "Работа"

    # Условие 2: числовое значение из первого столбца должно встречаться в первом столбце первого Treeview
    if file1 is not None:
        # Извлечение всех чисел из первого столбца первого Treeview
        def extract_numbers(value):
            if pd.isna(value):
                return []
            value_str = str(value).replace(" ", "").split(",")
            return [num for num in value_str if re.match(r'^\d+$', str(num).strip())]

        report_numbers = set()
        for value in file1.iloc[:, 0]:
            report_numbers.update(extract_numbers(value))

        # Проверка, что значение первого столбца содержится в report_numbers
        def check_inclusion(value):
            if pd.isna(value):
                return False
            numbers = extract_numbers(value)
            return any(num in report_numbers for num in numbers)

        condition2 = file2.iloc[:, 0].apply(check_inclusion)
    else:
        condition2 = pd.Series([False] * len(file2))  # Если первый Treeview пуст, все строки исключаются по условию 2

    # Применение фильтров
    file2 = file2[condition1 & condition2]
    display_dataframe(file2, tree2, file1_path if file1_path else "Файл не загружен", row_counter_label2)

def main():
    # Инициализация переменных
    file1, file2 = None, None
    file1_path, file2_path = None, None
    original_file2 = None  # Сохранение исходных данных второго файла для сброса фильтров

    # Создание меток для счётчиков строк
    row_counter_label1 = ttk.Label(root, text="Строк: 0")
    row_counter_label2 = ttk.Label(root, text="Строк: 0")

    def load_file(file_num):
        nonlocal file1, file2, file1_path, file2_path, original_file2
        df, path = load_excel_file()
        if df is not None:
            if file_num == 1:
                df = apply_filter_report(df)  # Применение фильтра для отчёта
                file1, file1_path = df, path
                display_dataframe(file1, tree1, file1_path, row_counter_label1, is_first_tree=True)
                messagebox.showinfo("Ништяк", f"Загружен первый файл: {os.path.basename(path)}")
            else:
                df = apply_filter_pcn_data(df)  # Применение фильтра для данных ПЦН
                file2, file2_path = df, path
                original_file2 = df.copy()  # Сохранение исходных данных
                display_dataframe(file2, tree2, file2_path, row_counter_label2, is_first_tree=False)
                messagebox.showinfo("Ништяк", f"Загружен второй файл: {os.path.basename(path)}")

    # Макет интерфейса
    # Секция для первого файла
    ttk.Button(root, text="Закинуть отчет", command=lambda: load_file(1)).pack(pady=5)
    ttk.Button(root, text="Экспорт в Excel", command=lambda: export_to_excel(tree1, file1_path)).pack(pady=5, expand=True)
    frame1 = ttk.Frame(root)
    frame1.pack(pady=5, fill="both", expand=True)
    tree1 = ttk.Treeview(frame1, height=10)
    tree1.pack(side="left", fill="both", expand=True)
    scrollbar1 = ttk.Scrollbar(frame1, orient="vertical", command=tree1.yview)
    scrollbar1.pack(side="right", fill="y")
    tree1.configure(yscrollcommand=scrollbar1.set)
    display_dataframe(None, tree1, "Данные не загружены", row_counter_label1, is_first_tree=True)  # Начальный вызов
    row_counter_label1.pack(pady=2)

    # Секция для второго файла
    ttk.Button(root, text="Закинуть данные ПЦН", command=lambda: load_file(2)).pack(pady=5)
    ttk.Button(root, text="Экспорт в Excel", command=lambda: export_to_excel(tree2, file1_path, file2_path)).pack(pady=5, expand=True)

    frame_buttons = ttk.Frame(root)
    frame_buttons.pack(pady=5)
    ttk.Button(frame_buttons, text="НаеБАЛИ ! В работе без договора",
               command=lambda: filter_in_progress_no_contract(file2, tree2, row_counter_label2, file1, file1_path,
                                                              original_file2)).pack(side=ttk.LEFT, padx=5)
    ttk.Button(frame_buttons, text="ПроеБАЛИ ! Отключено с договором",
               command=lambda: filter_deactivated_with_contract(file2, tree2, row_counter_label2, file1, file1_path,
                                                                original_file2)).pack(side=ttk.LEFT, padx=5)
    frame2 = ttk.Frame(root)
    frame2.pack(pady=5, fill="both", expand=True)
    tree2 = ttk.Treeview(frame2, height=15)
    tree2.pack(side="left", fill="both", expand=True)
    scrollbar2 = ttk.Scrollbar(frame2, orient="vertical", command=tree2.yview)
    scrollbar2.pack(side="right", fill="y")
    tree2.configure(yscrollcommand=scrollbar2.set)
    display_dataframe(None, tree2, "Файл не загружен", row_counter_label2, is_first_tree=False)  # Начальный вызов
    row_counter_label2.pack(pady=2)

    # Кнопка выхода
    ttk.Button(root, text="В ужасе съебать", command=root.quit).pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()