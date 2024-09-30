import os
import xlrd
import xlwt
from tkinter import Tk, Button, messagebox, Entry, Label
from tkinter.filedialog import askdirectory

def process_excel_files(directory):
    output_filename = 'Подсчет таблиц.xls'
    output_book = xlwt.Workbook()
    output_sheet = output_book.add_sheet('Рассчет')

    current_col = 2
    row = 2

    max_table_name_length = 0  # Переменная для хранения максимальной длины названия таблицы

    for filename in os.listdir(directory):
        if filename.endswith('.xls'):
            file_path = os.path.join(directory, filename)
            try:
                workbook = xlrd.open_workbook(file_path)
                table_name = os.path.splitext(filename)[0]  # Получаем имя таблицы без расширения
                max_table_name_length = max(max_table_name_length, len(table_name))  # Обновляем максимальную длину

                has_counter = False  # Флаг для определения наличия значения "counter"

                for sheet in workbook.sheets():
                    try:
                        cell_value = find_counter_cell(sheet)
                        if cell_value is not None:
                            if not has_counter:
                                output_sheet.write(row, current_col, table_name)  # Записываем имя таблицы в первую строку
                                has_counter = True
                            output_sheet.write(row + 1, current_col, sheet.name)  # Исправлено: увеличиваем row на 1
                            output_sheet.write(row + 1, current_col + 1, cell_value)  # Исправлено: увеличиваем row на 1
                            row += 1  # Увеличиваем row на 1 для следующей строки
                    except:
                        continue

                if has_counter:
                    current_col += 4  # Увеличиваем current_col на 2, так как записываем 2 столбца (sheet.name и cell_value)
                    row = 2
            except:
                continue

    # Устанавливаем ширину колонки с названиями таблиц
    max_col = output_sheet.last_used_col

    for col_index in range(max_col + 1):
        if col_index % 2 == 0:  # Проверяем, является ли индекс колонки четным числом
            output_sheet.col(
                col_index).width = max_table_name_length * 256  # Умножаем на 256, чтобы получить ширину в единицах 1/256 символа

    try:
        output_book.save(output_filename)
        messagebox.showinfo("Готово", f"Результаты записаны в файл {output_filename}")
    except:
        messagebox.showerror("Ошибка", "Не удалось записать результаты. Проверьте, закрыт ли у вас старый файл Подсчет таблиц.xls")

def find_counter_cell(sheet):
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            cell_value = sheet.cell_value(row, col)
            if isinstance(cell_value, str) and cell_value.lower() == 'counter':
                if row > 0:
                    return sheet.cell_value(row - 1, col)
                else:
                    return None

    return None

def select_directory():
    directory = askdirectory(title='Выберите папку с рабочими таблицами .xls')
    if directory:
        directory_entry.delete(0, "end")
        directory_entry.insert(0, directory)

def run_program():
    directory = directory_entry.get()
    if directory:
        process_excel_files(directory)
    else:
        messagebox.showerror("Ошибка", "Папка не выбрана.")

# Создание главного окна
root = Tk()
root.title("Подсчет рабочих таблиц")

# Создание виджетов
directory_label = Label(root, text="Папка:")
directory_entry = Entry(root, width=50)
select_button = Button(root, text="Выбрать папку", command=select_directory)
run_button = Button(root, text="Запуск", command=run_program)
close_button = Button(root, text="Закрыть", command=root.quit)

# Размещение виджетов на главном окне
directory_label.grid(row=0, column=0, sticky="e")
directory_entry.grid(row=0, column=1, padx=5, pady=5, columnspan=2)
select_button.grid(row=0, column=3, padx=5, pady=5)
run_button.grid(row=1, column=1, padx=5, pady=10)
close_button.grid(row=1, column=2, padx=5, pady=10)

# Установка параметров размещения кнопок в центре окна
root.grid_rowconfigure(2, weight=1)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(3, weight=1)

# Запуск главного цикла программы
root.mainloop()
