import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
import os
import sys


def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


EXCEL_FILE = os.path.join(get_base_dir(), "detailing_records.xlsx")

SERVICES = [
    "Химчистка салона",
    "Полировка кузова",
    "Мойка двигателя",
    "Детейлинг полный",
    "Покрытие керамикой",
    "Защитная пленка",
    "Чернение резины",
    "Озонирование",
    "Обезжиривание",
    "Другое",
]

HEADERS = [
    "#", "Марка", "г.номер", "Владелец",
    "Услуга", "Цена", "Материал",
    "Сумма рабочего", "Расход", "Касса", "Примечание"
]


class DetailingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Детейлинг — Учёт клиентов")
        self.root.minsize(460, 520)
        self.root.resizable(True, True)
        self.root.configure(bg="#f5f5f5")

        self._build_ui()

    def _build_ui(self):
        # Header
        header = tk.Frame(self.root, bg="#1a1a2e", pady=12)
        header.pack(fill="x")
        tk.Label(
            header, text="🚗  Детейлинг — Учёт клиентов",
            font=("Helvetica", 15, "bold"),
            fg="white", bg="#1a1a2e"
        ).pack()

        # Main form frame
        form_frame = tk.Frame(self.root, bg="#f5f5f5", padx=24, pady=16)
        form_frame.pack(fill="both", expand=True)

        self.entries = {}

        def add_label_entry(row, label, key, readonly=False, hint=""):
            tk.Label(
                form_frame, text=label, anchor="w",
                font=("Helvetica", 11), bg="#f5f5f5", fg="#333"
            ).grid(row=row, column=0, sticky="w", pady=5, padx=(0, 12))

            if readonly:
                var = tk.StringVar(value="0")
                widget = tk.Label(
                    form_frame, textvariable=var,
                    font=("Helvetica", 11, "bold"),
                    bg="#e8f4e8", fg="#2d6a2d",
                    relief="groove", width=22, anchor="w", padx=6
                )
                widget.grid(row=row, column=1, sticky="ew", pady=5)
                self.entries[key] = var
            else:
                var = tk.StringVar()
                widget = tk.Entry(
                    form_frame, textvariable=var,
                    font=("Helvetica", 11),
                    relief="solid", bd=1, bg="white"
                )
                widget.grid(row=row, column=1, sticky="ew", pady=5)
                if hint:
                    widget.insert(0, hint)
                    widget.config(fg="grey")
                    widget.bind("<FocusIn>", lambda e, w=widget, h=hint: self._clear_hint(e, w, h))
                    widget.bind("<FocusOut>", lambda e, w=widget, h=hint: self._restore_hint(e, w, h))
                self.entries[key] = var

        def add_label_combo(row, label, key):
            tk.Label(
                form_frame, text=label, anchor="w",
                font=("Helvetica", 11), bg="#f5f5f5", fg="#333"
            ).grid(row=row, column=0, sticky="w", pady=5, padx=(0, 12))

            var = tk.StringVar()
            combo = ttk.Combobox(
                form_frame, textvariable=var,
                values=SERVICES, state="readonly",
                font=("Helvetica", 11), width=22
            )
            combo.grid(row=row, column=1, sticky="ew", pady=5)
            self.entries[key] = var

        form_frame.columnconfigure(1, weight=1)

        add_label_entry(0,  "Марка *",       "marka")
        add_label_entry(1,  "г.номер *",     "gnomer")
        add_label_entry(2,  "Владелец",      "owner",    hint="87XXXXXXXXX")
        add_label_combo(3,  "Услуга *",      "service")
        add_label_entry(4,  "Цена (₸) *",   "price")
        add_label_entry(5,  "Материал (₸)", "material")
        add_label_entry(6,  "Сумма рабочего", "worker",  readonly=True)
        add_label_entry(7,  "Расход (₸)",   "rashod")
        add_label_entry(8,  "Касса (₸)",    "kassa",    readonly=True)
        add_label_entry(9,  "Примечание",   "note")

        # Live calculation bindings
        self.entries["price"].trace_add("write", self._recalc)
        self.entries["material"].trace_add("write", self._recalc)
        self.entries["rashod"].trace_add("write", self._recalc)

        # Divider
        tk.Frame(self.root, height=1, bg="#ddd").pack(fill="x", padx=24)

        # Save button
        btn_frame = tk.Frame(self.root, bg="#f5f5f5", pady=14)
        btn_frame.pack()
        tk.Button(
            btn_frame,
            text="💾  Сохранить",
            font=("Helvetica", 12, "bold"),
            bg="#1a1a2e", fg="white",
            activebackground="#16213e",
            activeforeground="white",
            relief="flat", padx=28, pady=8,
            cursor="hand2",
            command=self._save
        ).pack()

        tk.Label(
            self.root,
            text="* — обязательные поля",
            font=("Helvetica", 9), fg="#999", bg="#f5f5f5"
        ).pack(pady=(0, 8))

    def _clear_hint(self, event, widget, hint):
        if widget.get() == hint:
            widget.delete(0, tk.END)
            widget.config(fg="black")

    def _restore_hint(self, event, widget, hint):
        if widget.get() == "":
            widget.insert(0, hint)
            widget.config(fg="grey")

    def _safe_float(self, key):
        try:
            val = self.entries[key].get().strip()
            if val in ("", "87XXXXXXXXX"):
                return 0.0
            return float(val)
        except ValueError:
            return 0.0

    def _recalc(self, *args):
        price = self._safe_float("price")
        material = self._safe_float("material")
        rashod = self._safe_float("rashod")
        self.entries["worker"].set(str(round(price - material, 2)))
        self.entries["kassa"].set(str(round(price - rashod, 2)))

    def _get_next_number(self):
        if not os.path.exists(EXCEL_FILE):
            return 1
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        return ws.max_row  # row 1 = header, so max_row - 1 data rows + 1 = max_row

    def _ensure_excel(self):
        if not os.path.exists(EXCEL_FILE):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(HEADERS)
            # Style header
            from openpyxl.styles import Font, PatternFill, Alignment
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1a1a2e")
                cell.alignment = Alignment(horizontal="center")
            wb.save(EXCEL_FILE)

    def _save(self):
        marka   = self.entries["marka"].get().strip()
        gnomer  = self.entries["gnomer"].get().strip()
        service = self.entries["service"].get().strip()
        price_s = self.entries["price"].get().strip()

        # Validation
        missing = []
        if not marka:   missing.append("Марка")
        if not gnomer:  missing.append("г.номер")
        if not service: missing.append("Услуга")
        if not price_s: missing.append("Цена")

        if missing:
            messagebox.showerror(
                "Ошибка",
                f"❌ Заполните обязательные поля: {', '.join(missing)}"
            )
            return

        try:
            price = float(price_s)
        except ValueError:
            messagebox.showerror("Ошибка", "❌ Цена должна быть числом")
            return

        owner    = self.entries["owner"].get().strip()
        if owner == "87XXXXXXXXX":
            owner = ""
        material = self._safe_float("material")
        rashod   = self._safe_float("rashod")
        worker   = round(price - material, 2)
        kassa    = round(price - rashod, 2)
        note     = self.entries["note"].get().strip()

        self._ensure_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        num = ws.max_row  # header is row 1, so next num = max_row
        row_num = f"#{num}"
        ws.append([row_num, marka, gnomer, owner, service,
                   price, material, worker, rashod, kassa, note])
        wb.save(EXCEL_FILE)

        messagebox.showinfo("Сохранено", f"✅ Запись {row_num} сохранена")
        self._clear_form()

    def _clear_form(self):
        skip_readonly = {"worker", "kassa"}
        hints = {"owner": "87XXXXXXXXX"}
        for key, var in self.entries.items():
            if key in skip_readonly:
                var.set("0")
            else:
                var.set("")
        # Restore hints visually
        for widget in self.root.winfo_children():
            self._restore_hints_recursive(widget, hints)

    def _restore_hints_recursive(self, widget, hints):
        if isinstance(widget, tk.Entry):
            pass  # hints handled by FocusOut bindings
        for child in widget.winfo_children():
            self._restore_hints_recursive(child, hints)


def main():
    root = tk.Tk()
    app = DetailingApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
