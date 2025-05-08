import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog


def update_excel_file():
    root = tk.Tk()
    root.withdraw()

    try:
        # 1. Загрузка файлов
        messagebox.showinfo("Выбор файла", "Выберите файл с ID (источник данных)")
        file_with_id = filedialog.askopenfilename(
            title="Выберите файл с ID",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not file_with_id:
            return

        messagebox.showinfo("Выбор файла", "Выберите файл для обновления (без ID)")
        file_to_update = filedialog.askopenfilename(
            title="Выберите файл для обновления",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not file_to_update:
            return

        # 2. Чтение файлов
        df_source = pd.read_excel(file_with_id)  # Файл с ID
        df_target = pd.read_excel(file_to_update)  # Файл для обновления

        # 3. Определение столбцов для сопоставления
        key_columns = simpledialog.askstring(
            "Ключевые столбцы",
            "Введите названия столбцов для сопоставления (через запятую, например 'ФИО,дата_p'):",
            initialvalue="ФИО,дата_p"
        )
        if not key_columns:
            return

        key_columns = [col.strip() for col in key_columns.split(',')]

        # 4. Проверка наличия столбцов
        for col in key_columns + ['id']:
            if col not in df_source.columns:
                messagebox.showerror("Ошибка", f"Столбец '{col}' не найден в файле с ID!")
                return
            if col not in df_target.columns and col != 'id':
                messagebox.showerror("Ошибка", f"Столбец '{col}' не найден в файле для обновления!")
                return

        # 5. Создаем словарь {ключ: id} из исходного файла
        id_mapping = df_source.set_index(key_columns)['id'].to_dict()

        # 6. Обновляем ID в целевом файле
        def get_id(row):
            key = tuple(row[col] for col in key_columns)
            return id_mapping.get(key)

        df_target['id'] = df_target.apply(get_id, axis=1)

        # 7. Сохраняем изменения прямо в исходный файл (с подтверждением)
        confirm = messagebox.askyesno(
            "Подтверждение",
            f"Вы уверены, что хотите обновить файл?\n{file_to_update}\n\n"
            "Рекомендуется сделать backup перед продолжением."
        )
        if not confirm:
            return

        # 8. Сохранение (перезапись исходного файла)
        with pd.ExcelWriter(file_to_update, engine='openpyxl', mode='w') as writer:
            df_target.to_excel(writer, index=False)

        messagebox.showinfo(
            "Успех",
            f"Файл успешно обновлен!\n\n"
            f"Обновлено записей: {len(df_target[df_target['id'].notna()])}\n"
            f"Всего записей: {len(df_target)}\n\n"
            f"Файл: {file_to_update}"
        )

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")


if __name__ == "__main__":
    update_excel_file()