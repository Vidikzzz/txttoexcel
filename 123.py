import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def parse_fixed_width_file(input_file_path, output_file_path):
    """
    Парсит текстовый файл фиксированной ширины (формат F15444025) и сохраняет в Excel.
    
    Параметры:
        input_file_path (str): Путь к входному текстовому файлу.
        output_file_path (str): Путь к выходному Excel-файлу.
    """
    # Определяем позиции столбцов на основе фиксированной ширины
    colspecs = [
        (0, 12),    # КОД
        (12, 52),   # НАИМЕНОВАНИЕ МАТЕРИАЛА
        (52, 59),   # ЕДИН.
        (59, 72),   # СРЕДНЕВЗВЕШ.
        (72, 86),   # ФАКТ С
        (86, 100),  # СТОИМОСТЬ С
        (100, 114), # ФАКТ
        (114, 128)  # СТОИМОСТЬ
    ]
    
    try:
        # Читаем файл с фиксированной шириной столбцов
        df = pd.read_fwf(input_file_path, colspecs=colspecs, header=0, encoding='utf-8')
        
        # Удаляем возможные лишние пробелы в заголовках и данных
        df.columns = df.columns.str.strip()
        df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
        
        # Сохраняем в Excel
        df.to_excel(output_file_path, index=False, engine='openpyxl')
        messagebox.showinfo("Успех", f"Данные успешно сохранены в:\n{output_file_path}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}\n\nУбедитесь, что файл соответствует формату F15444025.")

def select_input_file():
    """Открывает диалоговое окно для выбора входного файла."""
    file_path = filedialog.askopenfilename(
        title="Выберите текстовый файл (формат F15444025)",
        filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")]
    )
    if file_path:
        input_entry.delete(0, tk.END)
        input_entry.insert(0, file_path)

def select_output_file():
    """Открывает диалоговое окно для выбора выходного файла."""
    file_path = filedialog.asksaveasfilename(
        title="Сохранить как Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")]
    )
    if file_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, file_path)

def run_conversion():
    """Запускает конвертацию выбранного файла."""
    input_file = input_entry.get()
    output_file = output_entry.get()
    
    if not input_file or not output_file:
        messagebox.showwarning("Предупреждение", "Пожалуйста, укажите входной и выходной файлы!")
        return
    
    parse_fixed_width_file(input_file, output_file)

# Создаем графический интерфейс
root = tk.Tk()
root.title("Конвертер F15444025 в Excel")

# Входной файл
tk.Label(root, text="Входной файл (формат F15444025):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
input_entry = tk.Entry(root, width=50)
input_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Выбрать...", command=select_input_file).grid(row=0, column=2, padx=5, pady=5)

# Выходной файл
tk.Label(root, text="Выходной Excel файл:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=5, pady=5)
tk.Button(root, text="Выбрать...", command=select_output_file).grid(row=1, column=2, padx=5, pady=5)

# Кнопка запуска
tk.Button(root, text="Конвертировать", command=run_conversion).grid(row=2, column=1, pady=10)

# Информация о формате
info_label = tk.Label(
    root,
    text="Поддерживается только формат F15444025.\nПожалуйста, загружайте корректные файлы.",
    fg="gray"
)
info_label.grid(row=3, column=0, columnspan=3, pady=5)

root.mainloop()