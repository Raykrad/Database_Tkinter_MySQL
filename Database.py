from tkinter import *
from tkinter import Frame, ttk, simpledialog, messagebox, filedialog
from tkinter.filedialog import askopenfilename, asksaveasfilename
from PIL import Image, ImageTk, ImageDraw, ImageFont, UnidentifiedImageError
import pandas as pd
from docx import Document
import mysql.connector
import json
import os
import xml.etree.ElementTree as ET

db = mysql.connector.connect(
        host = "localhost",
        user = "root",
        passwd = "123123123",
        database = "ui"
    )

cursor = db.cursor(buffered=True)

#cursor.execute("CREATE DATABASE UI")

#cursor.execute("ALTER TABLE licenses ADD COLUMN image_path TEXT")

cursor.execute("CREATE TABLE IF NOT EXISTS software ( software_id INT AUTO_INCREMENT PRIMARY KEY, \
    name VARCHAR(100), \
    version VARCHAR(100), \
    developer VARCHAR(100), \
    license_expiration VARCHAR(12))")

cursor.execute("CREATE TABLE IF NOT EXISTS users ( user_id INT AUTO_INCREMENT PRIMARY KEY, \
    first_name VARCHAR(100), \
    last_name VARCHAR(100), \
    role VARCHAR(100), \
    password VARCHAR(100))")

cursor.execute("CREATE TABLE IF NOT EXISTS licenses (license_id INT AUTO_INCREMENT PRIMARY KEY, \
    software_id INT, \
    user_id INT, \
    license_number varchar(100), \
    status varchar(50), \
    FOREIGN KEY (software_id) REFERENCES software(software_id), \
    FOREIGN KEY (user_id) REFERENCES users(user_id))")

def center_window(window):
    window.after(10, lambda: _center_window(window))

def _center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    window.geometry(f"+{x}+{y}")

def create_oval_image(text, font_size, padding=10, bg_color="black", text_color="white"):
    font = ImageFont.truetype("arial.ttf", font_size)
    text_bbox = font.getbbox(text)
    button_size = (text_bbox[2] - text_bbox[0] + padding * 2, text_bbox[3] - text_bbox[1] + padding * 2)

    image = Image.new("RGBA", button_size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((0, 0, *button_size), radius=button_size[1] // 2, fill=bg_color)
    draw.text(((button_size[0] - text_bbox[2]) // 2, (button_size[1] - text_bbox[3]) // 2), text, font=font, fill=text_color)

    return ImageTk.PhotoImage(image)

def all_destroy():
    window.destroy()
    if all_licenses:
        all_licenses.destroy()
    if delete_window:
        delete_window.destroy()
    if add_window:
        add_window.destroy()
    if update_window:
        update_window.destroy()

window = Tk()
window.title("UI")
center_window(window)
window.iconbitmap('c:/gui/ray.ico')
window.geometry("600x400")
window.configure(background='black')

background = Image.open('c:/gui/i.jpg')
bg_image = ImageTk.PhotoImage(background)
bg_label = Label(window, image=bg_image)
bg_label.place(relwidth=1, relheight=1)

bg_canvas = Canvas(window, width=512, height=512, highlightthickness=0)
bg_canvas.pack(fill="both", expand = True)

def resizer(e):
    global bg1, resized_bg, new_bg
    bg1 = Image.open('c:/gui/i.jpg')
    resized_bg = bg1.resize((e.width, e.height), Image.LANCZOS)
    new_bg = ImageTk.PhotoImage(resized_bg)
    bg_canvas.create_image(0, 0, image=new_bg, anchor='nw')

class DBHelper:
    def delete_licenses():
        def delete():
            try:
                license_id = int(license_id_entry.get())

                cursor.execute("SELECT software_id, user_id FROM licenses WHERE license_id = %s", (license_id,))
                result = cursor.fetchone()
                if result is None:
                    messagebox.showerror("Ошибка", f"Лицензия с ID {license_id} не найдена.")
                    return
                software_id, user_id = result

                cursor.execute("SELECT COUNT(*) FROM licenses WHERE software_id = %s", (software_id,))
                software_count = cursor.fetchone()[0]
                cursor.execute("SELECT COUNT(*) FROM licenses WHERE user_id = %s", (user_id,))
                user_count = cursor.fetchone()[0]

                if software_count == 1 and user_count == 1: # Удаляем только если ключи используются только в этой записи
                    cursor.execute("DELETE FROM licenses WHERE license_id = %s", (license_id,))
                    cursor.execute("DELETE FROM software WHERE software_id = %s", (software_id,))
                    cursor.execute("DELETE FROM users WHERE user_id = %s", (user_id,))
                elif software_count == 1:
                    cursor.execute("DELETE FROM licenses WHERE license_id = %s", (license_id,))
                    cursor.execute("DELETE FROM software WHERE software_id = %s", (software_id,))
                    messagebox.showinfo("Успех", "Лицензия и программное обеспечение успешно удалены.")
                elif user_count == 1:
                    cursor.execute("DELETE FROM licenses WHERE license_id = %s", (license_id,))
                    cursor.execute("DELETE FROM users WHERE user_id = %s", (user_id,))
                    messagebox.showinfo("Успех", "Лицензия и пользователь успешно удалены.")
                else:
                    cursor.execute("DELETE FROM licenses WHERE license_id = %s", (license_id,))
                    messagebox.showinfo("Успех", "Лицензия успешно удалена (связанные данные не удалены).")
                db.commit()
                messagebox.showinfo("Успех", "Лицензия успешно удалена.")
                refresh_tree_delayed()
                delete_window.destroy()
            except ValueError:
                messagebox.showerror("Ошибка", "Введите корректное число.")
            except mysql.connector.Error as err:
                db.rollback()
                messagebox.showerror("Ошибка", f"Ошибка работы с базой данных: {err}")
            except Exception as e:
                db.rollback()
                messagebox.showerror("Ошибка", f"Произошла непредвиденная ошибка: {e}")

        global delete_window
        delete_window = Toplevel()
        delete_window.title("Удаление")
        center_window(delete_window)
        delete_window.iconbitmap('c:/gui/ray.ico')
        delete_window.geometry("200x100")
        delete_window.configure(bg='black')

        Label(delete_window, text="Введите ID лицензии", bg='black', fg='white').place(relx=0.5, rely=0.1, anchor="center")

        license_id_entry = Entry(delete_window)
        license_id_entry.place(relx=0.5, rely=0.4, anchor="center")

        Button(delete_window, text="Удалить", command=delete).place(relx=0.5, rely=0.8, anchor="center")
    
    def add_licenses():
        def add():
            data = {}
            try:
                for label_text, var_name in field_data:
                    if var_name == "status":
                        data[var_name] = status_combo.get()
                    else:
                        value = entry_widgets[var_name].get()
                        if value == "":
                            raise ValueError(f"Поле '{label_text}' не заполнено.")
                        if label_text in ["ID лицензии", "ID ПО", "ID пользователя"]:
                            data[var_name] = int(value)
                        else:
                            data[var_name] = value

                if any(v == "" for v in data.values() if v != "Неактивна" and v != "Активна"):
                    raise ValueError("Заполните все поля.")

                cursor.execute("INSERT INTO software (software_id) VALUES (%s)", (data['software_id'],))
                cursor.execute("INSERT INTO users (user_id) VALUES (%s)", (data['user_id'],))
                cursor.execute("INSERT INTO licenses (license_id, software_id, user_id, license_number, status) VALUES (%s, %s, %s, %s, %s)", tuple(data.values()))
                db.commit()
                messagebox.showinfo("Успех", "Лицензия успешно добавлена.")
                refresh_tree_delayed()
                add_window.destroy()

            except mysql.connector.Error as err:
                messagebox.showerror("Ошибка", f"Ошибка базы данных: {err}\nПараметры: {data}")
            except ValueError as e:
                messagebox.showerror("Ошибка", str(e))
            except Exception as e:
                messagebox.showerror("Ошибка", f"Произошла неизвестная ошибка: {e}")

        global add_window
        add_window = Toplevel()
        add_window.title("Добавление")
        center_window(add_window)
        add_window.iconbitmap('c:/gui/ray.ico')
        add_window.geometry("290x200")
        add_window.configure(bg='black')

        field_data = [
            ("ID лицензии", "license_id"),
            ("ID ПО", "software_id"),
            ("ID пользователя", "user_id"),
            ("Номер лицензии", "license_number"),
            ("Статус лицензии", "status"),
        ]

        entry_widgets = {}
        row = 0
        for label_text, var_name in field_data:
            label = Label(add_window, text=label_text, bg='black', fg='white')
            label.grid(row=row, column=0, sticky="w", padx=10, pady=5)
            if var_name != "status":
                entry = Entry(add_window)
                entry.grid(row=row, column=1, sticky="e", padx=10, pady=5)
                entry_widgets[var_name] = entry
            row += 1

        status_combo = ttk.Combobox(add_window, values=["Активна", "Неактивна"], state="readonly")
        status_combo.grid(row=row-1, column=1, sticky="e", padx=10, pady=5)
        status_combo.set("Активна")
        row +=1

        Button(add_window, text="Добавить", command=add).grid(row=row, column=0, columnspan=2, pady=10)
    
    def update_licenses():
        def get_license_id():
            license_id = simpledialog.askinteger("ID Лицензии", "Введите ID лицензии для редактирования:")
            if license_id is not None:
                show_update_form(license_id)

        def show_update_form(license_id):
            def update():
                data = {}
                try:
                    for label_text, var_name in field_data:
                        if var_name == "status":
                            data[var_name] = status_combo.get()
                        else:
                            value = entry_widgets[var_name].get()
                            if value == "" and var_name not in ["software_id", "user_id"]:
                                raise ValueError(f"Поле '{label_text}' не заполнено.")
                            elif label_text in ["ID лицензии", "ID ПО", "ID пользователя"]:
                                data[var_name] = int(value) if value else None
                            else:
                                data[var_name] = value

                    if not data.get('license_id'):
                        raise ValueError("ID лицензии обязательно для обновления.")

                    if data.get('software_id') is not None:
                        cursor.execute("SELECT COUNT(*) FROM software WHERE software_id = %s", (data['software_id'],))
                        if cursor.fetchone()[0] == 0:
                            raise ValueError(f"ПО с ID {data['software_id']} не существует.")
                    if data.get('user_id') is not None:
                        cursor.execute("SELECT COUNT(*) FROM users WHERE user_id = %s", (data['user_id'],))
                        if cursor.fetchone()[0] == 0:
                            raise ValueError(f"Пользователь с ID {data['user_id']} не существует.")

                    sql_command = "UPDATE licenses SET "
                    update_fields = []
                    values = []
                    for key, value in data.items():
                        if key != "license_id" and value is not None:
                            update_fields.append(f"{key} = %s")
                            values.append(value)

                    if update_fields:
                        sql_command += ", ".join(update_fields) + " WHERE license_id = %s"
                        values.append(license_id)
                        cursor.execute(sql_command, tuple(values))
                        db.commit()
                        messagebox.showinfo("Успех", "Лицензия успешно обновлена.")
                        refresh_tree_delayed()
                        update_window.destroy()
                    else:
                        messagebox.showinfo("Информация", "Нет данных для обновления.")

                except mysql.connector.Error as err:
                    messagebox.showerror("Ошибка", f"Ошибка базы данных: {err}\nПараметры: {data}")
                except ValueError as e:
                    messagebox.showerror("Ошибка", str(e))
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Произошла неизвестная ошибка: {e}")

            global update_window
            update_window = Toplevel()
            update_window.title("Обновление лицензии")
            center_window(update_window)
            update_window.iconbitmap('c:/gui/ray.ico')
            update_window.geometry("290x200")
            update_window.configure(bg='black')

            field_data = [
                ("ID лицензии", "license_id"),
                ("ID ПО", "software_id"),
                ("ID пользователя", "user_id"),
                ("Номер лицензии", "license_number"),
                ("Статус лицензии", "status"),
            ]

            try:
                cursor.execute("SELECT * FROM licenses WHERE license_id = %s", (license_id,))
                license_data = cursor.fetchone()
                if license_data is None:
                    messagebox.showerror("Ошибка", f"Лицензия с ID {license_id} не найдена.")
                    update_window.destroy()
                    return

                entry_widgets = {}
                row = 0
                for i, (label_text, var_name) in enumerate(field_data):
                    label = Label(update_window, text=label_text, bg='black', fg='white')
                    label.grid(row=row, column=0, sticky="w", padx=10, pady=5)
                    if var_name != "status":
                        entry = Entry(update_window)
                        entry.insert(0, str(license_data[i]) if license_data[i] is not None else "")
                        entry.grid(row=row, column=1, sticky="e", padx=10, pady=5)
                        entry_widgets[var_name] = entry
                    row += 1

                status_combo = ttk.Combobox(update_window, values=["Активна", "Неактивна"], state="readonly")
                status_combo.grid(row=row-1, column=1, sticky="e", padx=10, pady=5)
                status_combo.set(license_data[4])
                row += 1

                Button(update_window, text="Обновить", command=update).grid(row=row, column=0, columnspan=2, pady=10)

            except mysql.connector.Error as err:
                messagebox.showerror("Ошибка", f"Ошибка базы данных: {err}")
                update_window.destroy()

        get_license_id()
        refresh_tree_delayed()

    def show_licenses():
        
        def search_records():
            tree.delete(*tree.get_children())
            search_value = search_entry.get()
            search_query = "SELECT * FROM licenses WHERE license_number LIKE %s OR status LIKE %s"
            try:
                cursor.execute(search_query, (f'%{search_value}%', f'%{search_value}%'))
                results = cursor.fetchall()
                for row in results:
                    tree.insert("", "end", values=row)
            except mysql.connector.Error as err:
                messagebox.showerror("Database Error", f"An error occurred: {err}")

        def export_to_excel():
            file_path = asksaveasfilename(defaultextension=".xlsx", 
                                            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                                            title="Сохранить как")
            if not file_path:
                return

            rows = [tree.item(row)["values"] for row in tree.get_children()]

            df = pd.DataFrame(rows, columns=columns)
            df.to_excel(file_path, sheet_name='Sheet1', index=False)
            print(f"Данные успешно экспортированы в файл: {file_path}")

        def import_from_excel():
            file_path = filedialog.askopenfilename(defaultextension=".xlsx", 
                                                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                                                    title="Выберите файл Excel")
            if not file_path:
                return

            try:
                df = pd.read_excel(file_path)

                required_columns = set(columns)
                imported_columns = set(df.columns)
                if required_columns != imported_columns:
                    messagebox.showerror("Ошибка", f"Несоответствие столбцов в файле Excel. Требуются столбцы: {required_columns}, а есть: {imported_columns}")
                    return

                df.fillna("", inplace=True)
                
                for col in ['license_id', 'software_id', 'user_id']:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

                for index, row in df.iterrows():
                    print(f"Обработка строки {index}: {row}")
                    try:
                        license_id = row['license_id']
                        software_id = row['software_id']
                        user_id = row['user_id']

                        cursor.execute("SELECT 1 FROM licenses WHERE license_id = %s", (license_id,))
                        exists_license = cursor.fetchone()
                        if exists_license:
                            print(f"Запись с license_id {license_id} уже существует, пропускаем.")
                            continue

                        cursor.execute("INSERT INTO software (software_id) VALUES (%s) ON DUPLICATE KEY UPDATE software_id = %s", (software_id, software_id))
                        cursor.execute("INSERT INTO users (user_id) VALUES (%s) ON DUPLICATE KEY UPDATE user_id = %s", (user_id, user_id))
                        
                        values = tuple(row.values)
                        sql = f"INSERT INTO licenses ({', '.join(columns)}) VALUES ({', '.join(['%s'] * len(columns))})"
                        print(f"SQL: {sql}, Values: {values}")
                        cursor.execute(sql, values)
                    except Exception as e:
                        print(f"Ошибка в строке {index}: {e}")
                        raise

                db.commit()
                refresh_tree()
                messagebox.showinfo("Успешно", f"Данные успешно импортированы из файла: {file_path}")

            except FileNotFoundError:
                messagebox.showerror("Ошибка", f"Файл не найден: {file_path}")
            except pd.errors.EmptyDataError:
                messagebox.showerror("Ошибка", f"Файл Excel пуст: {file_path}")
            except pd.errors.ParserError:
                messagebox.showerror("Ошибка", f"Ошибка при парсинге файла Excel: {file_path}")
            except mysql.connector.Error as err:
                db.rollback()
                messagebox.showerror("Ошибка", f"Ошибка записи в базу данных: {err}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Произошла непредвиденная ошибка: {e}")

        def get_data_from_tree(tree_view):
            rows = [tree_view.item(row)["values"] for row in tree_view.get_children()]
            return rows

        def update_tree_view(tree_view, data):
            tree_view.delete(*tree_view.get_children())
            for row in data:
                tree_view.insert("", "end", values=row)

        def export_to_xml(data):
            file_path = asksaveasfilename(defaultextension=".xml", filetypes=[("XML files", "*.xml"), ("All files", "*.*")], title="Сохранить как")
            if not file_path:
                return

            root = ET.Element("licenses")
            for row in data:
                license = ET.SubElement(root, "license")
                for i, col in enumerate(columns):
                    license.set(col, str(row[i]))

            tree_xml = ET.ElementTree(root)
            tree_xml.write(file_path, encoding="utf-8", xml_declaration=True)
            print(f"Данные успешно экспортированы в файл: {file_path}")

        def import_from_xml(tree_view):
            file_path = filedialog.askopenfilename(defaultextension=".xml", filetypes=[("XML files", "*.xml"), ("All files", "*.*")], title="Выберите файл XML")
            if not file_path:
                return

            try:
                tree = ET.parse(file_path)
                root = tree.getroot()
                imported_data = []

                for license in root.findall("license"):
                    row = []
                    for col in columns:
                        value = license.get(col)
                        if value is None:
                            messagebox.showerror("Ошибка", f"Отсутствует атрибут '{col}' в элементе 'license'.")
                            return
                        try:
                            if col == "license_id":
                                row.append(int(value))
                            else:
                                row.append(value)
                        except ValueError:
                            messagebox.showerror("Ошибка", f"Неверный формат данных в атрибуте '{col}'.")
                            return
                    imported_data.append(row)

                for row in imported_data:
                    license_id = row[columns.index('license_id')]
                    software_id = row[columns.index('software_id')]
                    user_id = row[columns.index('user_id')]

                    cursor.execute("SELECT 1 FROM licenses WHERE license_id = %s", (license_id,))
                    exists_license = cursor.fetchone()
                    if exists_license:
                        print(f"Запись с license_id {license_id} уже существует, пропускаем.")
                        continue

                    cursor.execute("INSERT INTO software (software_id) VALUES (%s) ON DUPLICATE KEY UPDATE software_id = %s", (software_id, software_id))
                    cursor.execute("INSERT INTO users (user_id) VALUES (%s) ON DUPLICATE KEY UPDATE user_id = %s", (user_id, user_id))
                    values = tuple(row)
                    sql = f"INSERT INTO licenses ({', '.join(columns)}) VALUES ({', '.join(['%s'] * len(columns))})"
                    cursor.execute(sql, values)

                db.commit()
                update_tree_view(tree_view, imported_data)
                messagebox.showinfo("Успешно", f"Данные успешно импортированы из файла: {file_path}")

            except FileNotFoundError:
                messagebox.showerror("Ошибка", f"Файл не найден: {file_path}")
            except ET.ParseError:
                messagebox.showerror("Ошибка", f"Ошибка при парсинге файла XML: {file_path}")
            except mysql.connector.Error as err:
                db.rollback()
                messagebox.showerror("Ошибка", f"Ошибка записи в базу данных: {err}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Произошла непредвиденная ошибка: {e}")

        def export_to_word():
            file_path = asksaveasfilename(defaultextension=".docx", 
                                            filetypes=[("Word files", "*.docx"), ("All files", "*.*")],
                                            title="Сохранить как")
            if not file_path:
                return

            rows = [tree.item(row)["values"] for row in tree.get_children()]

            document = Document()

            table = document.add_table(rows=len(rows) + 1, cols=len(columns))

            hdr_cells = table.rows[0].cells
            for i, col in enumerate(columns):
                hdr_cells[i].text = str(col)

            for row_num, row_data in enumerate(rows):
                row_cells = table.rows[row_num + 1].cells
                for col_num, cell_data in enumerate(row_data):
                    row_cells[col_num].text = str(cell_data)

            document.save(file_path)
            print(f"Данные успешно экспортированы в файл: {file_path}")

        def export_to_json():
            file_path = asksaveasfilename(defaultextension=".json", 
                                                    filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                                                    title="Сохранить как JSON")
            if not file_path:
                return

            rows = [tree.item(row)["values"] for row in tree.get_children()]
            df = pd.DataFrame(rows, columns=columns)
            try:
                data = df.to_dict(orient='records')
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
                print(f"Данные успешно экспортированы в JSON: {file_path}")
            except Exception as e:
                print(f"Ошибка экспорта JSON: {e}")

        def export_data(format, tree_view):
            data = get_data_from_tree(tree_view)
            if format == "Excel":
                export_to_excel()
            elif format == "Word":
                export_to_word()
            elif format == "JSON":
                export_to_json()
            elif format == "XML":
                export_to_xml(data)

        def load_config(filepath):
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return config
            except (FileNotFoundError, json.JSONDecodeError) as e:
                print(f"Ошибка загрузки конфигурации: {e}")
                return None

        def save_config(filepath, config):
            try:
                with open(filepath, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=4, ensure_ascii=False)
                print(f"Конфигурация сохранена: {filepath}")
            except Exception as e:
                print(f"Ошибка сохранения конфигурации: {e}")

        def apply_config(tree, config):
            if config:
                columns_config = config.get('columns', {})
                if columns_config:
                    for col in list(tree['columns']):
                        if col not in columns_config:
                            tree.heading(col, text="")
                            tree.column(col, width=0, stretch=False)
                            tree.column(col, anchor='center')
                            tree.forget(col)

                    for col, col_config in columns_config.items():
                        if col not in tree['columns']:
                            tree.column(col, width=col_config.get('width', 100), stretch=col_config.get('stretch', False), anchor='center')
                            tree.heading(col, text=col_config.get('heading', col))
                        else:
                            tree.heading(col, text=col_config.get('heading', col))
                            tree.column(col, width=col_config.get('width', 100), stretch=col_config.get('stretch', False), anchor='center')

                else:
                    print("Ошибка: Список столбцов пуст в конфигурации.")
            else:
                print("Ошибка: Конфигурация не загружена.")

        def load_table_config():
            filepath = askopenfilename(filetypes=[("JSON Files", "*.json")], title="Загрузить конфигурацию")
            if filepath:
                config = load_config(filepath)
                apply_config(tree, config)

        def save_table_config():
            filepath = filedialog.asksaveasfilename(filetypes=[("JSON Files", "*.json")], title="Сохранить конфигурацию")
            if filepath:
                if not filepath.lower().endswith(".json"):
                    result = messagebox.askyesno("Предупреждение", "Имя файла не содержит расширения '.json'. Добавить?")
                    if result:
                        filepath += ".json"
                    else:
                        return
                config = {
                    'columns': {col: {'heading': tree.heading(col), 'width': tree.column(col, option='width'), 'stretch': tree.column(col, option='stretch')} for col in tree['columns']}
                }
                save_config(filepath, config)

        def show_image(image_path):
            if not image_path or not os.path.exists(image_path):
                messagebox.showerror("Ошибка", "Изображение не найдено.")
                return

            try:
                img = Image.open(image_path)
                img.thumbnail((800, 600))
                photo = ImageTk.PhotoImage(img)
                image_window = Toplevel()
                image_window.title("Изображение")
                label = Label(image_window, image=photo)
                label.image = photo
                label.pack()
                image_window.mainloop()
            except FileNotFoundError:
                messagebox.showerror("Ошибка", f"Файл изображения не найден: {image_path}")
            except IOError as e:
                messagebox.showerror("Ошибка", f"Ошибка ввода-вывода: {e}")
            except UnidentifiedImageError:
                messagebox.showerror("Ошибка", f"Невозможно открыть изображение: {image_path}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Непредвиденная ошибка: {e}")

        def add_image(item_id):
            filepath = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif")], title="Выберите изображение")
            if filepath:
                try:
                    if item_id not in tree.get_children():
                        messagebox.showerror("Ошибка", "Элемент не найден в Treeview.")
                        return
                    license_id = tree.item(item_id)['values'][0]
                    sql = "UPDATE licenses SET image_path = %s WHERE license_id = %s"
                    val = (filepath, license_id)
                    cursor.execute(sql, val)
                    db.commit()
                    tree.set(item_id, "image_path", filepath)
                    messagebox.showinfo("Успешно", f"Изображение добавлено: {filepath}")
                    show_image(filepath)
                except mysql.connector.Error as err:
                    messagebox.showerror("Ошибка", f"Ошибка сохранения изображения в базе данных: {err}")
                    db.rollback()
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Общая ошибка: {e}")

        def on_double_click(event):
            try:
                item_id = tree.selection()[0]
                item = tree.item(item_id)
                image_path = item['values'][-1]

                if image_path and os.path.exists(image_path):
                    show_image(image_path)
                else:
                    if messagebox.askyesno("Изображение не найдено", "Хотите добавить изображение?"):
                        add_image(item_id)
            except IndexError:
                messagebox.showerror("Ошибка", "Не выбран элемент в Treeview.")
            except KeyError:
                messagebox.showerror("Ошибка", "В элементе Treeview отсутствует путь к изображению.")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Непредвиденная ошибка: {e}")

        window.withdraw()

        global all_licenses
        all_licenses = Toplevel()
        all_licenses.title("Лицензии")
        center_window(all_licenses)
        all_licenses.iconbitmap('c:/gui/ray.ico')
        all_licenses.geometry("900x600")
        all_licenses.configure(background='black')
        
        left_panel = Frame(all_licenses, bg='black', borderwidth=3, highlightthickness=2, highlightbackground='grey')
        left_panel.pack(side=LEFT, fill=Y, padx=10)

        add_button = Button(left_panel, text="Добавить лицензию", command=lambda: DBHelper.add_licenses(), borderwidth=3, highlightthickness=2, highlightbackground='grey')
        add_button.grid(row=0, column=0, padx=5, pady=15, sticky="ew")

        update_button = Button(left_panel, text="Обновить лицензию", command=lambda: DBHelper.update_licenses(), borderwidth=3, highlightthickness=2, highlightbackground='grey')
        update_button.grid(row=1, column=0, padx=5, pady=15, sticky="ew")

        delete_button = Button(left_panel, text="Удалить лицензию", command=lambda: DBHelper.delete_licenses(), borderwidth=3, highlightthickness=2, highlightbackground='grey')
        delete_button.grid(row=2, column=0, padx=5, pady=15, sticky="ew")

        excel_import_button = Button(left_panel, text="Импорт из Excel", command=import_from_excel, borderwidth=3, highlightthickness=2, highlightbackground='grey')
        excel_import_button.grid(row=3, column=0, padx=5, pady=15, sticky="ew")

        xml_import_button = Button(left_panel, text="Импорт из XML", command=lambda: import_from_xml(tree), borderwidth=3, highlightthickness=2, highlightbackground='grey')
        xml_import_button.grid(row=4, column=0, padx=5, pady=15, sticky="ew")

        load_config_button = Button(left_panel, text="Загрузить конфигурацию", command=load_table_config, borderwidth=3, highlightthickness=2, highlightbackground='grey')
        load_config_button.grid(row=5, column=0, padx=5, pady=15, sticky="ew")

        save_config_button = Button(left_panel, text="Сохранить конфигурацию", command=save_table_config, borderwidth=3, highlightthickness=2, highlightbackground='grey')
        save_config_button.grid(row=6, column=0, padx=5, pady=15, sticky="ew")

        export_combobox = ttk.Combobox(left_panel, values=["Excel", "Word", "JSON", "XML"], state="readonly")
        export_combobox.set("Выберите формат")
        export_combobox.grid(row=7, column=0, padx=5, pady=15, sticky="ew")

        export_button = ttk.Button(left_panel, text="Экспорт", command=lambda: export_data(export_combobox.get(), tree))
        export_button.grid(row=8, column=0, padx=5, pady=0, sticky="ew")

        search_frame = Frame(all_licenses, bg='black')
        search_frame.pack(pady=10)

        search_entry = Entry(search_frame, width=30)
        search_entry.pack(side=LEFT, padx=10)

        search_button = Button(search_frame, text="Поиск", command=search_records)
        search_button.pack(side=LEFT)

        cursor.execute("SELECT * FROM licenses")
        result = cursor.fetchall()

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", background="black", foreground="white", fieldbackground="black", rowheight=25)
        style.configure("Treeview.Heading", background="black", foreground="white", relief="flat")
        style.map('Treeview', background=[('selected', 'grey')], foreground=[('selected', 'white')])

        columns = [desc[0] for desc in cursor.description]

        tree = ttk.Treeview(all_licenses, columns=columns, show='headings', style="Treeview")

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor='center')

        for row in result:
            tree.insert("", END, values=row)

        tree.pack(expand=True, fill='both')

        tree.bind("<Double-1>", on_double_click)

        def refresh_tree():
            cursor.execute("SELECT * FROM licenses")
            result = cursor.fetchall()
            for item in tree.get_children():
                tree.delete(item)
            for row in result:
                tree.insert("", END, values=row)

        global refresh_tree_delayed
        def refresh_tree_delayed():
            all_licenses.after(50, refresh_tree)

        def on_closing():
            all_licenses.destroy()
            window.deiconify()

        all_licenses.protocol("WM_DELETE_WINDOW", on_closing)
    
label = Label(window, text='Учёт лицензий', bg='black', fg='white', font=("MS Reference Sans Serif", 18))
label.place(relx=0.33, rely=0.02, anchor="ne")

exit_button_img = create_oval_image("Выйти", font_size=16, bg_color="RGB(26,26,26)", text_color="white")
exit_button = Button(window, image=exit_button_img, command=all_destroy, bg='black', activebackground='black', borderwidth=0, highlightthickness=2,)
exit_button.place(relx=1, rely=0, anchor="ne")

show_licenses_img = create_oval_image("Показать все лицензии", font_size=16, bg_color="RGB(36, 36, 36)", text_color="white")
show_licenses_btn = Button(window, image=show_licenses_img, command=lambda: DBHelper.show_licenses(), bg='black', activebackground='black', borderwidth=0, highlightthickness=0)
show_licenses_btn.place(relx=0.5, rely=0.22, anchor="center")

delete_license_img = create_oval_image("Удалить лицензию", font_size=16, bg_color="RGB(36, 36, 36)", text_color="white")
delete_license_btn = Button(window, image=delete_license_img, command=lambda: DBHelper.delete_licenses(), bg='black', activebackground='black', borderwidth=0, highlightthickness=0)
delete_license_btn.place(relx=0.5, rely=0.37, anchor="center")

add_license_img = create_oval_image("Добавить лицензию", font_size=16, bg_color="RGB(36, 36, 36)", text_color="white")
add_license_btn = Button(window, image=add_license_img, command=lambda: DBHelper.add_licenses(), bg='black', activebackground='black', borderwidth=0, highlightthickness=0)
add_license_btn.place(relx=0.5, rely=0.52, anchor="center")

update_license_img = create_oval_image("Обновить лицензию", font_size=16, bg_color="RGB(36, 36, 36)", text_color="white")
update_license_btn = Button(window, image=update_license_img, command=lambda: DBHelper.update_licenses(), bg='black', activebackground='black', borderwidth=0, highlightthickness=0)
update_license_btn.place(relx=0.5, rely=0.67, anchor="center")

window.bind('<Configure>', resizer)
window.mainloop()
