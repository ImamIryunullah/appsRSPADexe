import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import os

DB_PATH = os.path.join("database", "mapping.db")

class ViewMappingScreen(tk.Frame):
    def __init__(self, master):
        super().__init__(master)

        self.label_title = ttk.Label(self, text="Data Mapping", font=("Arial", 16))
        self.label_title.pack(pady=10)

        # ================= SEARCH BAR =====================
        search_frame = tk.Frame(self)
        search_frame.pack(fill="x", pady=5)

        tk.Label(search_frame, text="Cari:").pack(side="left", padx=5)
        self.search_entry = tk.Entry(search_frame, width=30)
        self.search_entry.pack(side="left")

        self.search_button = ttk.Button(search_frame, text="Search", command=self.search_data)
        self.search_button.pack(side="left", padx=5)

        self.reset_button = ttk.Button(search_frame, text="Reset", command=self.load_table)
        self.reset_button.pack(side="left", padx=5)

        # ================= BUTTON ROW =====================
        button_frame = tk.Frame(self)
        button_frame.pack(fill="x", pady=5)

        ttk.Button(button_frame, text="Hapus Baris Terpilih", command=self.delete_row).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Hapus Semua Data", command=self.delete_all).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Refresh Data", command=self.load_table).pack(side="left", padx=5)

        # ================= TABLE AREA =====================
        table_frame = tk.Frame(self)
        table_frame.pack(fill="both", expand=True)

        self.scroll_y = tk.Scrollbar(table_frame, orient="vertical")
        self.scroll_y.pack(side="right", fill="y")

        self.scroll_x = tk.Scrollbar(table_frame, orient="horizontal")
        self.scroll_x.pack(side="bottom", fill="x")

        self.tree = ttk.Treeview(
            table_frame,
            yscrollcommand=self.scroll_y.set,
            xscrollcommand=self.scroll_x.set
        )
        self.tree.pack(fill="both", expand=True)

        self.scroll_y.config(command=self.tree.yview)
        self.scroll_x.config(command=self.tree.xview)

        self.load_table()

    # =====================================================
    # LOAD TABLE WITH AUTO NUMBERING
    # =====================================================
    def load_table(self):
        self.tree.delete(*self.tree.get_children())

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        cursor.execute("PRAGMA table_info(mapping)")
        columns_info = cursor.fetchall()

        if not columns_info:
            self.tree["columns"] = []
            self.tree.heading("#0", text="No data")
            conn.close()
            return

        db_columns = [col[1] for col in columns_info]
        db_columns.remove("id")

        # Add NO column for display only
        columns = db_columns

        self.tree["columns"] = columns
        self.tree.column("#0", width=0, stretch=tk.NO)

        # Setup NO column
        # self.tree.heading("no_urut", text="NO")
        # self.tree.column("no_urut", width=50, anchor="center")

        # Setup DB columns
        for col in db_columns:
            self.tree.heading(col, text=col.upper())
            self.tree.column(col, width=130, anchor="w")

        cursor.execute("SELECT * FROM mapping ORDER BY id ASC")
        rows = cursor.fetchall()

        no = 1
        for row in rows:
            row = row[1:]
            self.tree.insert("", "end", values=row)
            no += 1

        conn.close()

    # =====================================================
    # SEARCH FUNCTION
    # =====================================================
    def search_data(self):
        keyword = self.search_entry.get().strip()
        if keyword == "":
            messagebox.showinfo("Info", "Masukkan kata kunci pencarian.")
            return

        self.tree.delete(*self.tree.get_children())

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        cursor.execute("PRAGMA table_info(mapping)")
        columns_info = cursor.fetchall()
        db_columns = [col[1] for col in columns_info]

        query_parts = [f"{col} LIKE ?" for col in db_columns]
        params = [f"%{keyword}%"] * len(db_columns)

        sql = f"SELECT * FROM mapping WHERE {' OR '.join(query_parts)}"
        cursor.execute(sql, params)
        rows = cursor.fetchall()

        no = 1
        for row in rows:
            self.tree.insert("", "end", values=(no, *row))
            no += 1

        conn.close()

    # =====================================================
    # DELETE ONE ROW
    # =====================================================
    def delete_row(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Perhatian", "Pilih baris yang ingin dihapus!")
            return

        row_values = self.tree.item(selected[0])["values"]
        row_id = row_values[1]  # index 1 = id (karena 0 = NO urut)

        confirm = messagebox.askyesno("Konfirmasi", "Yakin ingin menghapus baris ini?")
        if not confirm:
            return

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM mapping WHERE id = ?", (row_id,))
        conn.commit()
        conn.close()

        self.load_table()
        messagebox.showinfo("Sukses", "Baris berhasil dihapus.")

    # =====================================================
    # DELETE ALL
    # =====================================================
    def delete_all(self):
        confirm = messagebox.askyesno("Konfirmasi", "Yakin ingin menghapus SEMUA data?")
        if not confirm:
            return

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM mapping")
        conn.commit()
        conn.close()

        self.load_table()
        messagebox.showinfo("Sukses", "Semua data berhasil dihapus.")
