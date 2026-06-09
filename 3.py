import os
import csv
import json
import sqlite3
import platform
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from docxtpl import DocxTemplate
import pymorphy3

from banks import (
    BanksDatabaseError,
    add_bank,
    delete_bank,
    get_bank_details,
    search_banks,
    update_bank,
)

FONT = ("Times New Roman", 12)
BG_OK = "white"
BG_ERROR = "#ffcccc"

TEMPLATES = {
    "zayavlenie": "templates/zayavlenie.docx",
    "creditors": "templates/creditors.docx",
    "property": "templates/property.docx",
}

OUTPUT_DIR = "output"
DATA_DIR = "data"
CLIENTS_FILE = os.path.join(DATA_DIR, "clients.json")

CODES_TXT_FILE = "codes_pass.txt"
CODES_DB_FILE = os.path.join(DATA_DIR, "passport_codes.db")

morph = pymorphy3.MorphAnalyzer()

fields = [
    ("Фамилия", "surname"),
    ("Имя", "name"),
    ("Отчество", "patronymic"),
    ("ФИО в родительном падеже", "fio_genitive"),
    ("Дата рождения", "birth_date"),
    ("Место рождения", "birth_place"),
    ("СНИЛС", "snils"),
    ("ИНН", "inn"),
    ("Серия паспорта", "passport_series"),
    ("Номер паспорта", "passport_number"),
    ("Дата выдачи паспорта", "passport_issue_date"),
    ("Код подразделения", "passport_code"),
    ("Кем выдан паспорт", "passport_issued_by"),
    ("Индекс", "postal_code"),
    ("Регион", "region"),
    ("Район", "district"),
    ("Город", "city"),
    ("Населённый пункт", "locality"),
    ("Улица", "street"),
    ("Дом", "house"),
    ("Корпус", "building"),
    ("Квартира", "apartment"),
    ("Доход", "income"),
]

manual_date_keys = {"birth_date", "passport_issue_date"}

digit_limits = {
    "snils": 11,
    "inn": 12,
    "passport_series": 4,
    "passport_number": 6,
    "postal_code": 6,
}

month_names = {
    "01": "января", "02": "февраля", "03": "марта",
    "04": "апреля", "05": "мая", "06": "июня",
    "07": "июля", "08": "августа", "09": "сентября",
    "10": "октября", "11": "ноября", "12": "декабря",
}


def ensure_files():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if not os.path.exists(CLIENTS_FILE):
        with open(CLIENTS_FILE, "w", encoding="utf-8") as file:
            json.dump([], file, ensure_ascii=False, indent=4)


def only_digits(value, max_len=None):
    digits = "".join(ch for ch in value if ch.isdigit())
    return digits[:max_len] if max_len else digits


def format_passport_code(value):
    digits = only_digits(value, 6)
    if len(digits) > 3:
        return f"{digits[:3]}-{digits[3:]}"
    return digits


def normalize_passport_code(value):
    digits = only_digits(value, 6)
    if len(digits) == 6:
        return f"{digits[:3]}-{digits[3:]}"
    return value.strip()


def load_clients():
    ensure_files()
    try:
        with open(CLIENTS_FILE, "r", encoding="utf-8") as file:
            return json.load(file)
    except Exception:
        return []


def save_clients(clients):
    ensure_files()
    with open(CLIENTS_FILE, "w", encoding="utf-8") as file:
        json.dump(clients, file, ensure_ascii=False, indent=4)


def recreate_passport_codes_db():
    ensure_files()

    if not os.path.exists(CODES_TXT_FILE):
        return

    if os.path.exists(CODES_DB_FILE):
        os.remove(CODES_DB_FILE)

    conn = sqlite3.connect(CODES_DB_FILE)
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE passport_codes (
            id TEXT,
            name TEXT NOT NULL,
            code TEXT NOT NULL UNIQUE,
            end_date TEXT
        )
    """)

    cursor.execute("CREATE INDEX idx_passport_codes_code ON passport_codes(code)")
    cursor.execute("CREATE INDEX idx_passport_codes_name ON passport_codes(name)")

    encodings = ["utf-8-sig", "utf-8", "cp1251"]
    opened_file = None

    for encoding in encodings:
        try:
            opened_file = open(CODES_TXT_FILE, "r", encoding=encoding, newline="")
            opened_file.read(100)
            opened_file.seek(0)
            break
        except UnicodeDecodeError:
            if opened_file:
                opened_file.close()
            opened_file = None

    if not opened_file:
        conn.close()
        return

    rows = []

    with opened_file as file:
        reader = csv.DictReader(file, delimiter=";")
        for row in reader:
            code = normalize_passport_code(row.get("CODE", ""))
            name = row.get("NAME", "").strip()
            record_id = row.get("ID", "").strip()
            end_date = row.get("END_DATE", "").strip()

            if code and name and len(only_digits(code)) == 6:
                rows.append((record_id, name, code, end_date))

    cursor.executemany("""
        INSERT OR REPLACE INTO passport_codes (id, name, code, end_date)
        VALUES (?, ?, ?, ?)
    """, rows)

    conn.commit()
    conn.close()


def create_passport_codes_db_if_needed():
    if not os.path.exists(CODES_DB_FILE):
        recreate_passport_codes_db()


def search_passport_codes(query, limit=10):
    create_passport_codes_db_if_needed()

    if not os.path.exists(CODES_DB_FILE):
        return []

    query = query.strip()
    digits = only_digits(query)

    conn = sqlite3.connect(CODES_DB_FILE)
    cursor = conn.cursor()

    if digits:
        like_value = f"{digits}%"
        cursor.execute("""
            SELECT code, name
            FROM passport_codes
            WHERE REPLACE(code, '-', '') LIKE ?
            ORDER BY code
            LIMIT ?
        """, (like_value, limit))
    else:
        like_value = f"%{query.upper()}%"
        cursor.execute("""
            SELECT code, name
            FROM passport_codes
            WHERE UPPER(name) LIKE ?
            ORDER BY code
            LIMIT ?
        """, (like_value, limit))

    rows = cursor.fetchall()
    conn.close()
    return rows


def count_codes_in_db():
    create_passport_codes_db_if_needed()

    if not os.path.exists(CODES_DB_FILE):
        return 0

    conn = sqlite3.connect(CODES_DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM passport_codes")
    count = cursor.fetchone()[0]
    conn.close()
    return count


def auto_date_format(event):
    entry = event.widget
    digits = only_digits(entry.get(), 8)

    result = ""
    for i, ch in enumerate(digits):
        if i == 2 or i == 4:
            result += "."
        result += ch

    entry.delete(0, tk.END)
    entry.insert(0, result)
    entry.configure(bg=BG_OK if len(result) == 10 else BG_ERROR)


def auto_digits_format(event, max_len):
    entry = event.widget
    digits = only_digits(entry.get(), max_len)
    entry.delete(0, tk.END)
    entry.insert(0, digits)
    entry.configure(bg=BG_OK if len(digits) == max_len else BG_ERROR)


def to_genitive_word(word):
    if not word:
        return ""

    parsed = morph.parse(word)[0]
    genitive = parsed.inflect({"gent"})

    return genitive.word.capitalize() if genitive else word


def fio_to_genitive(surname, name, patronymic):
    return " ".join(
        to_genitive_word(part)
        for part in [surname, name, patronymic]
        if part
    )


def split_date(date_text, prefix):
    result = {
        f"{prefix}_day": "",
        f"{prefix}_month_number": "",
        f"{prefix}_month": "",
        f"{prefix}_year": "",
    }

    parts = date_text.split(".")
    if len(parts) == 3:
        day, month, year = parts
        result[f"{prefix}_day"] = day
        result[f"{prefix}_month_number"] = month
        result[f"{prefix}_month"] = month_names.get(month, "")
        result[f"{prefix}_year"] = year

    return result


def safe_folder_name(name):
    bad_chars = '<>:"/\\|?*'
    for char in bad_chars:
        name = name.replace(char, "_")
    return name.strip()


def open_folder(path):
    system_name = platform.system()

    if system_name == "Windows":
        os.startfile(path)
    elif system_name == "Darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class YuristApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Yurist+")
        self.root.geometry("980x860")

        self.entries = {}
        self.clients = load_clients()
        self.codes_count = count_codes_in_db()
        self.current_code_results = []
        self.bank_window = None

        self.create_menu()
        self.create_ui()
        self.refresh_clients_list()

    def create_menu(self):
        menu_bar = tk.Menu(self.root)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Выход", command=self.root.quit)
        menu_bar.add_cascade(label="Файл", menu=file_menu)

        menu_bar.add_command(label="Банки", command=self.open_banks_window)

        self.root.config(menu=menu_bar)

    def open_banks_window(self):
        if self.bank_window and self.bank_window.window.winfo_exists():
            self.bank_window.window.lift()
            self.bank_window.window.focus_force()
            return

        self.bank_window = BanksWindow(self.root)

    def create_ui(self):
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        frame = ttk.Frame(canvas)

        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        tk.Label(
            frame,
            text="Yurist+ — карточка клиента",
            font=("Times New Roman", 16, "bold")
        ).grid(row=0, column=0, columnspan=3, pady=15)

        row = 1

        tk.Label(frame, text="Сохранённые клиенты", font=FONT).grid(
            row=row, column=0, sticky="w", padx=10, pady=4
        )

        self.client_combo = ttk.Combobox(frame, width=52, font=FONT, state="readonly")
        self.client_combo.grid(row=row, column=1, padx=10, pady=4)

        tk.Button(
            frame,
            text="Загрузить",
            font=FONT,
            command=self.load_selected_client
        ).grid(row=row, column=2, padx=10, pady=4)

        row += 1

        tk.Label(
            frame,
            text=f"База кодов DB: {self.codes_count} записей",
            font=FONT
        ).grid(row=row, column=0, columnspan=3, sticky="w", padx=10, pady=4)

        row += 1

        for label_text, key in fields:
            tk.Label(frame, text=label_text, font=FONT).grid(
                row=row, column=0, sticky="w", padx=10, pady=4
            )

            entry = tk.Entry(frame, width=55, font=FONT)
            entry.grid(row=row, column=1, padx=10, pady=4)

            if key in manual_date_keys:
                entry.bind("<KeyRelease>", auto_date_format)

            elif key == "passport_code":
                entry.bind("<KeyRelease>", self.on_passport_code_change)

            elif key in digit_limits:
                max_len = digit_limits[key]
                entry.bind(
                    "<KeyRelease>",
                    lambda event, limit=max_len: auto_digits_format(event, limit)
                )

            self.entries[key] = entry

            row += 1

            if key == "passport_code":
                self.code_results_box = tk.Listbox(frame, width=90, height=5, font=FONT)
                self.code_results_box.grid(row=row, column=0, columnspan=3, padx=10, pady=4)
                self.code_results_box.bind("<<ListboxSelect>>", self.select_passport_code)
                row += 1

        tk.Button(
            frame,
            text="Авто ФИО в родительном падеже",
            font=FONT,
            command=self.auto_fill_genitive
        ).grid(row=6, column=2, padx=10, pady=4)

        tk.Label(frame, text="Дата заполнения", font=FONT).grid(
            row=row, column=0, sticky="w", padx=10, pady=4
        )

        self.fill_date_entry = DateEntry(
            frame,
            width=20,
            font=FONT,
            date_pattern="dd.mm.yyyy",
            locale="ru_RU"
        )
        self.fill_date_entry.grid(row=row, column=1, sticky="w", padx=10, pady=4)

        row += 1

        tk.Button(
            frame,
            text="Сохранить клиента в базу",
            font=("Times New Roman", 12, "bold"),
            command=self.save_client_to_base,
            width=30
        ).grid(row=row, column=0, columnspan=3, pady=10)

        row += 1

        tk.Button(
            frame,
            text="Сформировать 3 документа",
            font=("Times New Roman", 12, "bold"),
            command=self.generate_documents,
            width=30
        ).grid(row=row, column=0, columnspan=3, pady=15)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def on_passport_code_change(self, event):
        entry = event.widget
        formatted = format_passport_code(entry.get())

        entry.delete(0, tk.END)
        entry.insert(0, formatted)

        digits = only_digits(formatted)

        if digits:
            entry.configure(bg=BG_OK if len(digits) == 6 else BG_ERROR)
            self.update_code_results(formatted)
        else:
            entry.configure(bg=BG_OK)
            self.code_results_box.delete(0, tk.END)

    def update_code_results(self, query):
        self.code_results_box.delete(0, tk.END)
        self.current_code_results = search_passport_codes(query, limit=10)

        for code, name in self.current_code_results:
            self.code_results_box.insert(tk.END, f"{code} — {name}")

    def select_passport_code(self, event):
        selection = self.code_results_box.curselection()

        if not selection:
            return

        index = selection[0]
        code, name = self.current_code_results[index]

        self.entries["passport_code"].delete(0, tk.END)
        self.entries["passport_code"].insert(0, code)
        self.entries["passport_code"].configure(bg=BG_OK)

        self.entries["passport_issued_by"].delete(0, tk.END)
        self.entries["passport_issued_by"].insert(0, name)
        self.entries["passport_issued_by"].configure(bg=BG_OK)

    def auto_fill_genitive(self):
        surname = self.entries["surname"].get().strip()
        name = self.entries["name"].get().strip()
        patronymic = self.entries["patronymic"].get().strip()

        genitive = fio_to_genitive(surname, name, patronymic)

        self.entries["fio_genitive"].delete(0, tk.END)
        self.entries["fio_genitive"].insert(0, genitive)

    def validate_fields(self):
        errors = []

        for key, max_len in digit_limits.items():
            value = self.entries[key].get().strip()

            if value and len(value) != max_len:
                self.entries[key].configure(bg=BG_ERROR)
                errors.append(key)
            else:
                self.entries[key].configure(bg=BG_OK)

        passport_code = self.entries["passport_code"].get().strip()
        if passport_code and len(only_digits(passport_code)) != 6:
            self.entries["passport_code"].configure(bg=BG_ERROR)
            errors.append("passport_code")
        else:
            self.entries["passport_code"].configure(bg=BG_OK)

        for key in manual_date_keys:
            value = self.entries[key].get().strip()

            if value and len(value) != 10:
                self.entries[key].configure(bg=BG_ERROR)
                errors.append(key)
            else:
                self.entries[key].configure(bg=BG_OK)

        return len(errors) == 0

    def collect_data(self):
        data = {}

        for key, entry in self.entries.items():
            data[key] = entry.get().strip()

        data["passport_code"] = normalize_passport_code(data["passport_code"])
        data["fill_date"] = self.fill_date_entry.get().strip()
        data["fio"] = f"{data['surname']} {data['name']} {data['patronymic']}".strip()

        if not data["fio_genitive"]:
            data["fio_genitive"] = fio_to_genitive(
                data["surname"],
                data["name"],
                data["patronymic"]
            )

        data.update(split_date(data["fill_date"], "fill"))
        data.update(split_date(data["birth_date"], "birth"))
        data.update(split_date(data["passport_issue_date"], "passport_issue"))

        data["full_address"] = (
            f"{data['postal_code']}, {data['region']}, {data['district']}, "
            f"{data['city']}, {data['locality']}, {data['street']}, "
            f"д. {data['house']}, корп. {data['building']}, кв. {data['apartment']}"
        )

        data["passport_full"] = (
            f"{data['passport_series']} {data['passport_number']}, "
            f"выдан: {data['passport_issued_by']} {data['passport_issue_date']}, "
            f"к/п {data['passport_code']}"
        )

        return data

    def fill_form(self, data):
        for key, entry in self.entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, data.get(key, ""))
            entry.configure(bg=BG_OK)

        if data.get("fill_date"):
            self.fill_date_entry.delete(0, tk.END)
            self.fill_date_entry.insert(0, data.get("fill_date"))

    def refresh_clients_list(self):
        self.clients = load_clients()

        names = []
        for client in self.clients:
            fio = client.get("fio", "").strip()
            if fio:
                names.append(fio)

        self.client_combo["values"] = names

    def save_client_to_base(self):
        if not self.validate_fields():
            messagebox.showerror("Ошибка", "Проверь красные поля.")
            return

        data = self.collect_data()

        if not data["surname"] or not data["name"]:
            messagebox.showerror("Ошибка", "Заполни хотя бы фамилию и имя клиента.")
            return

        self.clients = load_clients()
        updated = False

        for index, client in enumerate(self.clients):
            if client.get("fio") == data["fio"]:
                self.clients[index] = data
                updated = True
                break

        if not updated:
            self.clients.append(data)

        save_clients(self.clients)
        self.refresh_clients_list()
        messagebox.showinfo("Yurist+", "Клиент сохранён в базу.")

    def load_selected_client(self):
        selected_fio = self.client_combo.get()

        if not selected_fio:
            messagebox.showerror("Ошибка", "Выбери клиента из списка.")
            return

        self.clients = load_clients()

        for client in self.clients:
            if client.get("fio") == selected_fio:
                self.fill_form(client)
                messagebox.showinfo("Yurist+", "Клиент загружен.")
                return

    def generate_documents(self):
        if not self.validate_fields():
            messagebox.showerror("Ошибка", "Проверь красные поля.")
            return

        data = self.collect_data()
        self.save_client_to_base()

        client_folder_name = safe_folder_name(
            f"{data['surname']} {data['name']} {data['patronymic']}"
        )
        client_folder = os.path.join(OUTPUT_DIR, client_folder_name)
        os.makedirs(client_folder, exist_ok=True)

        for doc_name, template_path in TEMPLATES.items():
            if not os.path.exists(template_path):
                messagebox.showerror("Ошибка", f"Не найден шаблон: {template_path}")
                return

            doc = DocxTemplate(template_path)
            doc.render(data)

            output_path = os.path.join(
                client_folder,
                f"{doc_name}_{data['surname']}_{data['name']}.docx"
            )

            doc.save(output_path)

        open_folder(client_folder)
        messagebox.showinfo("Готово", "Документы сформированы и папка клиента открыта.")


class BanksWindow:
    SEARCH_PLACEHOLDER = "Поиск банка по названию, ИНН или ОГРН"

    def __init__(self, parent):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("Банки")
        self.window.geometry("900x600")
        self.window.minsize(760, 480)

        self.current_results = []
        self.search_placeholder_active = True

        self.create_ui()
        self.show_search_placeholder()
        self.refresh_results()

    def create_ui(self):
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(1, weight=1)

        top_frame = ttk.Frame(self.window, padding=(10, 10, 10, 6))
        top_frame.grid(row=0, column=0, sticky="ew")
        top_frame.columnconfigure(0, weight=1)

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(
            top_frame,
            textvariable=self.search_var,
            font=("Times New Roman", 14),
        )
        self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.search_entry.bind("<FocusIn>", self.on_search_focus_in)
        self.search_entry.bind("<FocusOut>", self.on_search_focus_out)
        self.search_entry.bind("<KeyRelease>", self.on_search_change)

        tk.Button(
            top_frame,
            text="Добавить банк",
            font=FONT,
            command=self.open_add_dialog,
            width=18,
        ).grid(row=0, column=1, sticky="e")

        list_frame = ttk.Frame(self.window, padding=(10, 0, 10, 10))
        list_frame.grid(row=1, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.banks_tree = ttk.Treeview(
            list_frame,
            columns=("bank",),
            show="headings",
            selectmode="browse",
        )
        self.banks_tree.heading("bank", text="Название банка | ИНН | ОГРН")
        self.banks_tree.column("bank", width=820, minwidth=500, stretch=True)
        self.banks_tree.grid(row=0, column=0, sticky="nsew")
        self.banks_tree.bind("<Double-Button-1>", self.open_details_for_selected)
        self.banks_tree.bind("<Button-3>", self.show_context_menu)
        self.banks_tree.bind("<Button-2>", self.show_context_menu)

        scrollbar = ttk.Scrollbar(
            list_frame,
            orient="vertical",
            command=self.banks_tree.yview,
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.banks_tree.configure(yscrollcommand=scrollbar.set)

        self.context_menu = tk.Menu(self.window, tearoff=0)
        self.context_menu.add_command(
            label="Изменить",
            command=self.open_edit_dialog_for_selected,
        )
        self.context_menu.add_command(label="Удалить", command=self.remove_selected_bank)

    def show_search_placeholder(self):
        self.search_placeholder_active = True
        self.search_var.set(self.SEARCH_PLACEHOLDER)
        self.search_entry.configure(fg="gray")

    def hide_search_placeholder(self):
        if self.search_placeholder_active:
            self.search_placeholder_active = False
            self.search_var.set("")
            self.search_entry.configure(fg="black")

    def on_search_focus_in(self, event):
        self.hide_search_placeholder()

    def on_search_focus_out(self, event):
        if not self.search_var.get().strip():
            self.show_search_placeholder()

    def on_search_change(self, event):
        if not self.search_placeholder_active:
            self.refresh_results()

    def get_search_query(self):
        if self.search_placeholder_active:
            return ""

        return self.search_var.get().strip()

    def refresh_results(self):
        for item_id in self.banks_tree.get_children():
            self.banks_tree.delete(item_id)

        try:
            self.current_results = search_banks(self.get_search_query())
        except BanksDatabaseError as exc:
            self.current_results = []
            messagebox.showerror("База банков", str(exc), parent=self.window)
            return

        for bank in self.current_results:
            self.banks_tree.insert(
                "",
                tk.END,
                iid=str(bank["id"]),
                values=(self.format_bank_row(bank),),
            )

    def format_bank_row(self, bank):
        return " | ".join((
            bank.get("name", ""),
            bank.get("inn", ""),
            bank.get("ogrn", ""),
        ))

    def get_selected_bank_id(self):
        selection = self.banks_tree.selection()
        if not selection:
            return None

        return int(selection[0])

    def select_row_under_pointer(self, event):
        row_id = self.banks_tree.identify_row(event.y)
        if row_id:
            self.banks_tree.selection_set(row_id)
            self.banks_tree.focus(row_id)
            return True

        return False

    def show_context_menu(self, event):
        if not self.select_row_under_pointer(event):
            return

        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def open_add_dialog(self):
        BankEditDialog(
            self.window,
            title="Добавить банк",
            on_save=self.add_bank_from_dialog,
        )

    def open_edit_dialog_for_selected(self):
        bank_id = self.get_selected_bank_id()
        if not bank_id:
            messagebox.showerror("База банков", "Выберите банк для изменения.", parent=self.window)
            return

        try:
            bank = get_bank_details(bank_id)
        except BanksDatabaseError as exc:
            messagebox.showerror("База банков", str(exc), parent=self.window)
            return

        if not bank:
            messagebox.showerror("База банков", "Банк не найден.", parent=self.window)
            self.refresh_results()
            return

        BankEditDialog(
            self.window,
            title="Изменить банк",
            bank=bank,
            on_save=lambda data: self.update_bank_from_dialog(bank_id, data),
        )

    def open_details_for_selected(self, event=None):
        bank_id = self.get_selected_bank_id()
        if not bank_id:
            return

        try:
            bank = get_bank_details(bank_id)
        except BanksDatabaseError as exc:
            messagebox.showerror("База банков", str(exc), parent=self.window)
            return

        if not bank:
            messagebox.showerror("База банков", "Банк не найден.", parent=self.window)
            self.refresh_results()
            return

        BankDetailsDialog(self.window, bank, on_edit=self.open_edit_dialog_for_selected)

    def add_bank_from_dialog(self, data):
        try:
            add_bank(data)
        except BanksDatabaseError as exc:
            messagebox.showerror("База банков", str(exc), parent=self.window)
            return False

        self.refresh_results()
        messagebox.showinfo("База банков", "Банк добавлен.", parent=self.window)
        return True

    def update_bank_from_dialog(self, bank_id, data):
        try:
            update_bank(bank_id, data)
        except BanksDatabaseError as exc:
            messagebox.showerror("База банков", str(exc), parent=self.window)
            return False

        self.refresh_results()
        tree_id = str(bank_id)
        if self.banks_tree.exists(tree_id):
            self.banks_tree.selection_set(tree_id)
            self.banks_tree.focus(tree_id)
        messagebox.showinfo("База банков", "Изменения банка сохранены.", parent=self.window)
        return True

    def remove_selected_bank(self):
        bank_id = self.get_selected_bank_id()
        if not bank_id:
            messagebox.showerror("База банков", "Выберите банк для удаления.", parent=self.window)
            return

        bank_name = self.banks_tree.item(str(bank_id), "values")[0].split(" | ", 1)[0]
        if not messagebox.askyesno(
            "Удаление банка",
            f"Удалить банк «{bank_name}» из базы данных?",
            parent=self.window,
        ):
            return

        try:
            delete_bank(bank_id)
        except BanksDatabaseError as exc:
            messagebox.showerror("База банков", str(exc), parent=self.window)
            return

        self.refresh_results()
        messagebox.showinfo("База банков", "Банк удалён.", parent=self.window)


class BankEditDialog:
    def __init__(self, parent, title, on_save, bank=None):
        self.parent = parent
        self.on_save = on_save
        self.bank = bank or {}
        self.fields = {}

        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("520x270")
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()

        self.create_ui()
        self.fill_fields()
        self.fields["name"].focus_set()

    def create_ui(self):
        form_frame = ttk.Frame(self.window, padding=12)
        form_frame.pack(fill="both", expand=True)
        form_frame.columnconfigure(1, weight=1)

        for row, (label_text, key) in enumerate((
            ("Название", "name"),
            ("ИНН", "inn"),
            ("ОГРН", "ogrn"),
            ("Адрес", "address"),
        )):
            tk.Label(form_frame, text=label_text, font=FONT).grid(
                row=row, column=0, sticky="w", padx=(0, 8), pady=6
            )
            entry = tk.Entry(form_frame, font=FONT)
            entry.grid(row=row, column=1, sticky="ew", pady=6)
            self.fields[key] = entry

        buttons_frame = ttk.Frame(form_frame)
        buttons_frame.grid(row=4, column=0, columnspan=2, sticky="e", pady=(14, 0))

        tk.Button(
            buttons_frame,
            text="Сохранить",
            font=FONT,
            command=self.save,
            width=14,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            buttons_frame,
            text="Отмена",
            font=FONT,
            command=self.window.destroy,
            width=14,
        ).pack(side="left")

    def fill_fields(self):
        for key, entry in self.fields.items():
            entry.insert(0, self.bank.get(key, ""))

    def collect_data(self):
        return {
            "name": self.fields["name"].get().strip(),
            "inn": self.fields["inn"].get().strip(),
            "ogrn": self.fields["ogrn"].get().strip(),
            "address": self.fields["address"].get().strip(),
        }

    def save(self):
        if self.on_save(self.collect_data()):
            self.window.destroy()


class BankDetailsDialog:
    def __init__(self, parent, bank, on_edit=None):
        self.parent = parent
        self.bank = bank
        self.on_edit = on_edit

        self.window = tk.Toplevel(parent)
        self.window.title("Информация о банке")
        self.window.geometry("560x420")
        self.window.minsize(460, 340)
        self.window.transient(parent)

        self.create_ui()

    def create_ui(self):
        frame = ttk.Frame(self.window, padding=12)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(4, weight=1)
        frame.rowconfigure(5, weight=1)

        self.add_value_row(frame, "Название", self.bank.get("name", ""), 0)
        self.add_value_row(frame, "ИНН", self.bank.get("inn", ""), 1)
        self.add_value_row(frame, "ОГРН", self.bank.get("ogrn", ""), 2)
        self.add_value_row(frame, "Адрес", self.bank.get("address", ""), 3)
        self.add_text_row(frame, "Филиалы", "\n".join(self.bank.get("branches", [])), 4)
        self.add_text_row(frame, "Альтернативные названия", "\n".join(self.bank.get("aliases", [])), 5)

        buttons_frame = ttk.Frame(frame)
        buttons_frame.grid(row=6, column=0, columnspan=2, sticky="e", pady=(10, 0))

        if self.on_edit:
            tk.Button(
                buttons_frame,
                text="Изменить",
                font=FONT,
                command=self.open_edit,
                width=14,
            ).pack(side="left", padx=(0, 8))

        tk.Button(
            buttons_frame,
            text="Закрыть",
            font=FONT,
            command=self.window.destroy,
            width=14,
        ).pack(side="left")

    def add_value_row(self, frame, label_text, value, row):
        tk.Label(frame, text=label_text, font=FONT).grid(
            row=row, column=0, sticky="nw", padx=(0, 8), pady=4
        )
        tk.Label(
            frame,
            text=value or "—",
            font=FONT,
            anchor="w",
            justify="left",
            wraplength=360,
        ).grid(row=row, column=1, sticky="ew", pady=4)

    def add_text_row(self, frame, label_text, value, row):
        tk.Label(frame, text=label_text, font=FONT).grid(
            row=row, column=0, sticky="nw", padx=(0, 8), pady=4
        )
        text = tk.Text(frame, height=4, font=FONT, wrap="word")
        text.grid(row=row, column=1, sticky="nsew", pady=4)
        text.insert("1.0", value or "—")
        text.configure(state="disabled")

    def open_edit(self):
        self.window.destroy()
        self.on_edit()


if __name__ == "__main__":
    ensure_files()
    create_passport_codes_db_if_needed()
    root = tk.Tk()
    app = YuristApp(root)
    root.mainloop()