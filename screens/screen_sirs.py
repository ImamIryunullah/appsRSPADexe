import customtkinter as ctk
from tkinter import messagebox, filedialog, ttk
import pandas as pd
import sqlite3
import threading

class SirsScreen(ctk.CTkFrame):
    def __init__(self, master, db_path="database/mapping.db"):
        super().__init__(master)
    
        self.db_path = db_path  # Path ke database SQLite
        self.file_path = None
        self.df_preview = None
        self.current_font_size = 11
        self.is_processing = False  # Flag untuk mencegah double click
        
        # === TITLE ===
        title = ctk.CTkLabel(self, text="Optimasi SIRS", font=("Arial", 20))
        title.pack(pady=10)
        
        # === TOP BAR (Upload + Search + Zoom) ===
        top_bar = ctk.CTkFrame(self)
        top_bar.pack(fill="x", padx=10)
        
        # Upload Button
        btn_upload = ctk.CTkButton(top_bar, text="Upload File Excel", width=150, command=self.upload_excel)
        btn_upload.pack(side="left", padx=5, pady=10)
        # Search field
        self.search_var = ctk.StringVar()
        entry_search = ctk.CTkEntry(top_bar, placeholder_text="Cari...", textvariable=self.search_var, width=200)
        entry_search.pack(side="left", padx=10)
        entry_search.bind("<KeyRelease>", self.filter_preview)
        # Zoom Buttons
        btn_zoom_in = ctk.CTkButton(top_bar, text="+", width=40, command=lambda: self.change_font_size(1))
        btn_zoom_in.pack(side="right", padx=5)
        btn_zoom_out = ctk.CTkButton(top_bar, text="-", width=40, command=lambda: self.change_font_size(-1))
        btn_zoom_out.pack(side="right")
        # File Info
        self.label_file = ctk.CTkLabel(self, text="Belum ada file yang dipilih")
        self.label_file.pack(pady=5)

        # === PREVIEW FRAME ===
        self.preview_frame = ctk.CTkFrame(self)
        self.preview_frame.pack(pady=10, fill="both", expand=True)

        # === RUN BUTTON ===
        btn_run = ctk.CTkButton(self, text="Mulai Optimasi", command=self.run_process)
        btn_run.pack(pady=15)
        
        # === PROGRESS BAR & STATUS ===
        self.progress_frame = ctk.CTkFrame(self)
        self.progress_frame.pack(pady=5, fill="x", padx=10)
        self.progress_frame.pack_forget()  # Sembunyikan dulu
        
        self.progress_label = ctk.CTkLabel(self.progress_frame, text="Memulai proses...")
        self.progress_label.pack(pady=5)
        
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, width=400)
        self.progress_bar.pack(pady=5)
        self.progress_bar.set(0)
        
        self.progress_detail = ctk.CTkLabel(self.progress_frame, text="", font=("Arial", 9))
        self.progress_detail.pack(pady=2)

        # === SAVE BUTTON ===
        btn_save = ctk.CTkButton(self, text="Simpan Hasil", command=self.save_result)
        btn_save.pack(pady=5)
        
        # === QUICK SAVE BUTTON ===
        btn_quick_save = ctk.CTkButton(self, text="Quick Save (Simple)", command=self.quick_save, fg_color="green")
        btn_quick_save.pack(pady=5)
        
        # === EXPORT PDF BUTTON ===
        btn_export_pdf = ctk.CTkButton(self, text="Export ke PDF", command=self.export_to_pdf, fg_color="orange")
        btn_export_pdf.pack(pady=5)

    def upload_excel(self):
        filetypes = [("Excel Files", "*.xlsx"), ("Excel Files", "*.xls")]
        filepath = filedialog.askopenfilename(title="Pilih File Excel", filetypes=filetypes)

        if not filepath:
            return

        self.file_path = filepath
        self.label_file.configure(text=f"File dipilih: {filepath}")

        try:
            self.df_preview = pd.read_excel(filepath)
            self.show_preview(self.df_preview)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca file Excel!\n\n{e}")

    def show_preview(self, df: pd.DataFrame):
        for widget in self.preview_frame.winfo_children():
            widget.destroy()

        container = ctk.CTkFrame(self.preview_frame)
        container.pack(fill="both", expand=True)

        scroll_y = ttk.Scrollbar(container, orient="vertical")
        scroll_y.pack(side="right", fill="y")

        scroll_x = ttk.Scrollbar(container, orient="horizontal")
        scroll_x.pack(side="bottom", fill="x")

        self.table = ttk.Treeview(container, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.table.pack(fill="both", expand=True)

        scroll_y.config(command=self.table.yview)
        scroll_x.config(command=self.table.xview)

        style = ttk.Style()
        style.configure("Treeview", font=("Arial", self.current_font_size))
        style.configure("Treeview.Heading", font=("Arial", self.current_font_size, "bold"))

        self.table["columns"] = list(df.columns)
        self.table["show"] = "headings"

        for col in df.columns:
            self.table.heading(col, text=col)
            self.table.column(col, width=150, anchor="w")

        for _, row in df.head(300).iterrows():
            self.table.insert("", "end", values=list(row))

    def filter_preview(self, event=None):
        if self.df_preview is None:
            return

        keyword = self.search_var.get().lower()

        for item in self.table.get_children():
            self.table.delete(item)

        df_filtered = self.df_preview[
            self.df_preview.apply(lambda row: row.astype(str).str.lower().str.contains(keyword).any(), axis=1)
        ]

        for _, row in df_filtered.head(300).iterrows():
            self.table.insert("", "end", values=list(row))

    def change_font_size(self, delta):
        self.current_font_size += delta
        self.current_font_size = max(self.current_font_size, 8)

        style = ttk.Style()
        style.configure("Treeview", font=("Arial", self.current_font_size))
        style.configure("Treeview.Heading", font=("Arial", self.current_font_size, "bold"))

    def get_sirs_column(self, usia_tahun, usia_bulan, usia_hari, gender):
        """Menentukan kolom SIRS berdasarkan usia dan jenis kelamin"""
        gender = str(gender).upper().strip()
        
        # Konversi ke integer, default 0
        try:
            usia_tahun = int(float(usia_tahun)) if pd.notna(usia_tahun) and str(usia_tahun).strip() != "" else 0
        except:
            usia_tahun = 0
            
        try:
            usia_bulan = int(float(usia_bulan)) if pd.notna(usia_bulan) and str(usia_bulan).strip() != "" else 0
        except:
            usia_bulan = 0
            
        try:
            usia_hari = int(float(usia_hari)) if pd.notna(usia_hari) and str(usia_hari).strip() != "" else 0
        except:
            usia_hari = 0

        # LOGIKA BERDASARKAN USIA TAHUN (prioritas tertinggi)
        if usia_tahun >= 85:
            return f"<85 thn_{gender}"
        
        if 80 <= usia_tahun <= 84:
            return f"80-84 thn_{gender}"
        
        if 75 <= usia_tahun <= 79:
            return f"75-79 thn_{gender}"
        
        if 70 <= usia_tahun <= 74:
            return f"70-74 thn_{gender}"
        
        if 65 <= usia_tahun <= 69:
            return f"65-69 thn_{gender}"
        
        if 60 <= usia_tahun <= 64:
            return f"60-64 thn_{gender}"
        
        if 55 <= usia_tahun <= 59:
            return f"55-59 thn_{gender}"
        
        if 50 <= usia_tahun <= 54:
            return f"50-54 thn_{gender}"
        
        if 45 <= usia_tahun <= 49:
            return f"45-49 thn_{gender}"
        
        if 40 <= usia_tahun <= 44:
            return f"40-44 thn_{gender}"
        
        if 35 <= usia_tahun <= 39:
            return f"35-39 thn_{gender}"
        
        if 30 <= usia_tahun <= 34:
            return f"30-34 thn_{gender}"
        
        if 25 <= usia_tahun <= 29:
            return f"25-29 thn_{gender}"
        
        if 20 <= usia_tahun <= 24:
            return f"20-24 thn_{gender}"
        
        if 15 <= usia_tahun <= 19:
            return f"15-19 thn_{gender}"
        
        if 10 <= usia_tahun <= 14:
            return f"10-14 thn_{gender}"
        
        if 5 <= usia_tahun <= 9:
            return f"5-9 thn_{gender}"
        
        if 1 <= usia_tahun <= 4:
            return f"1-4 thn_{gender}"

        # JIKA USIA TAHUN = 0, CEK USIA BULAN
        if usia_tahun == 0:
            # Kategori berdasarkan bulan (untuk bayi)
            if usia_bulan >= 12:
                # Kalau bulan >= 12, seharusnya masuk tahun, tapi jaga-jaga
                return f"6-11 bln_{gender}"
            
            if 6 <= usia_bulan <= 11:
                return f"6-11 bln_{gender}"
            
            if 3 <= usia_bulan <= 5:
                return f"3- <6 bln_{gender}"
            
            # Jika bulan < 3, cek hari
            if usia_bulan >= 1 or usia_hari >= 29:
                return f"29 hr- <30 bln_{gender}"
            
            # Kategori berdasarkan hari (untuk neonatus)
            if 8 <= usia_hari <= 28:
                return f"8-28 hr_{gender}"
            
            if 1 <= usia_hari <= 7:
                return f"1-7 hr_{gender}"
            
            if usia_hari == 0:
                # Bisa jadi < 1 hari (dalam jam)
                return f"<1 Jam_{gender}"

        # Default jika tidak masuk kategori
        return None

    def run_process(self):
        """Proses optimasi SIRS dengan membaca data dari database mapping"""
        
        if self.df_preview is None:
            messagebox.showwarning("Belum ada file", "Silakan upload file Excel terlebih dahulu.")
            return
        
        if self.is_processing:
            messagebox.showwarning("Sedang Proses", "Optimasi sedang berjalan, mohon tunggu...")
            return
        
        # Jalankan proses di thread terpisah agar UI tidak freeze
        thread = threading.Thread(target=self._run_process_thread, daemon=True)
        thread.start()
    
    def _run_process_thread(self):
        """Thread worker untuk proses optimasi"""
        self.is_processing = True
        
        # Tampilkan progress bar
        self.progress_frame.pack(pady=5, fill="x", padx=10)
        self.update_progress(0, "Memulai proses optimasi...")

        try:
            # Koneksi ke database
            self.update_progress(0.05, "Menghubungkan ke database...")
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Cek apakah tabel mapping ada
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='mapping'")
            table_exists = cursor.fetchone()
            
            if not table_exists:
                conn.close()
                self.show_error(
                    "Database Error", 
                    "Tabel 'mapping' tidak ditemukan di database!\n\n"
                    "Pastikan:\n"
                    "1. Database sudah dibuat\n"
                    "2. Tabel 'mapping' sudah ada\n"
                    f"3. Path database benar: {self.db_path}"
                )
                return
            
            # Ambil semua kolom dari tabel mapping
            self.update_progress(0.1, "Memeriksa struktur tabel...")
            cursor.execute("PRAGMA table_info(mapping)")
            columns_info = cursor.fetchall()
            column_names = [col[1] for col in columns_info]
            
            # Cari nama kolom yang sesuai
            def find_column(possible_names):
                for possible in possible_names:
                    for col in column_names:
                        if col.lower().replace(" ", "").replace("_", "") == possible.lower().replace(" ", "").replace("_", ""):
                            return col
                return None
            
            col_kode_icd = find_column(["kode_icd", "KODE ICD", "kode icd", "KODE_ICD"])
            col_kelamin = find_column(["kelamin", "KELAMIN", "jenis_kelamin", "JENIS KELAMIN"])
            col_usia_tahun = find_column(["usia_tahun", "USIA TAHUN", "usia_thn", "USIA_TAHUN"])
            col_usia_bulan = find_column(["usia_bulan", "USIA BULAN", "usia_bln", "USIA_BULAN"])
            col_usia_hari = find_column(["usia_hari", "USIA HARI", "usia_hr", "USIA_HARI"])
            col_alasan_pulang = find_column(["alasan_pulang", "ALASAN PULANG", "alasan pulang", "ALASAN_PULANG", "status_pulang", "STATUS PULANG"])
            
            # Validasi kolom yang dibutuhkan
            missing_cols = []
            if not col_kode_icd:
                missing_cols.append("KODE ICD")
            if not col_kelamin:
                missing_cols.append("KELAMIN")
            if not col_usia_tahun:
                missing_cols.append("USIA TAHUN")
            
            if missing_cols:
                conn.close()
                self.show_error(
                    "Kolom Tidak Ditemukan",
                    f"Kolom berikut tidak ditemukan di database:\n{', '.join(missing_cols)}\n\n"
                    f"Kolom yang tersedia:\n{', '.join(column_names)}"
                )
                return
            
            # Baca data dari database
            self.update_progress(0.15, "Membaca data dari database...")
            query = f"""
            SELECT "{col_kode_icd}", "{col_kelamin}", "{col_usia_tahun}", 
                   "{col_usia_bulan}", "{col_usia_hari}", "{col_alasan_pulang}"
            FROM mapping
            WHERE "{col_kode_icd}" IS NOT NULL AND "{col_kode_icd}" != ''
            """
            
            df_mapping = pd.read_sql_query(query, conn)
            conn.close()
            
            # Rename kolom
            df_mapping.columns = ['kode_icd', 'kelamin', 'usia_tahun', 'usia_bulan', 'usia_hari', 'alasan_pulang']

            if df_mapping.empty:
                self.show_warning("Data Kosong", "Tidak ada data di tabel mapping!")
                return

            total_rows_db = len(df_mapping)
            self.update_progress(0.2, f"Data mapping dimuat: {total_rows_db} baris")

            # Copy dataframe SIRS
            df_sirs = self.df_preview.copy()

            # Pastikan kolom 'Kode ICD' ada
            icd_col = None
            for col in df_sirs.columns:
                if col.lower().strip() == 'kode icd':
                    icd_col = col
                    break
            
            if icd_col is None:
                self.show_error("Error", "Kolom 'Kode ICD' tidak ditemukan di file Excel!")
                return

            # Definisi kolom SIRS
            SIRS_COLUMNS = [
                "<1 Jam_L", "<1 Jam_P", "1-23 jam_L", "1-23 jam_P",
                "1-7 hr_L", "1-7 hr_P", "8-28 hr_L", "8-28 hr_P",
                "29 hr- <30 bln_L", "29 hr- <30 bln_P",
                "3- <6 bln_L", "3- <6 bln_P", "6-11 bln_L", "6-11 bln_P",
                "1-4 thn_L", "1-4 thn_P", "5-9 thn_L", "5-9 thn_P",
                "10-14 thn_L", "10-14 thn_P", "15-19 thn_L", "15-19 thn_P",
                "20-24 thn_L", "20-24 thn_P", "25-29 thn_L", "25-29 thn_P",
                "30-34 thn_L", "30-34 thn_P", "35-39 thn_L", "35-39 thn_P",
                "40-44 thn_L", "40-44 thn_P", "45-49 thn_L", "45-49 thn_P",
                "50-54 thn_L", "50-54 thn_P", "55-59 thn_L", "55-59 thn_P",
                "60-64 thn_L", "60-64 thn_P", "65-69 thn_L", "65-69 thn_P",
                "70-74 thn_L", "70-74 thn_P", "75-79 thn_L", "75-79 thn_P",
                "80-84 thn_L", "80-84 thn_P", "<85 thn_L", "<85 thn_P"
            ]

            self.update_progress(0.25, "Menyiapkan kolom SIRS...")
            # Pastikan semua kolom SIRS ada, reset ke 0
            for col in SIRS_COLUMNS:
                if col not in df_sirs.columns:
                    df_sirs[col] = 0
                else:
                    df_sirs[col] = 0

            # Reset kolom jumlah jika ada
            jumlah_cols = [
                "Jumlah Pasien Keluar Hidup dan Mati Menurut Jenis Kelamin_L",
                "Jumlah Pasien Keluar Hidup dan Mati Menurut Jenis Kelamin_P",
                "Jumlah Pasien Keluar Hidup dan Mati Menurut Jenis Kelamin_TOTAL",
                "Jumlah Pasien Keluar Mati_L",
                "Jumlah Pasien Keluar Mati_P",
                "Jumlah Pasien Keluar Mati_TOTAL"
            ]
            for col in jumlah_cols:
                if col not in df_sirs.columns:
                    df_sirs[col] = 0
                else:
                    df_sirs[col] = 0

            # Proses setiap kode ICD di SIRS
            total_sirs = len(df_sirs)
            processed = 0
            
            self.update_progress(0.3, f"Memproses {total_sirs} kode ICD...")
            
            for idx, sirs_row in df_sirs.iterrows():
                processed += 1
                progress_pct = 0.3 + (0.6 * processed / total_sirs)  # 30% - 90%
                
                kode_icd_sirs = str(sirs_row[icd_col]).strip().upper()
                
                if kode_icd_sirs == "" or kode_icd_sirs == "NAN":
                    continue

                # Update progress setiap 10 baris
                if processed % 10 == 0 or processed == total_sirs:
                    self.update_progress(
                        progress_pct, 
                        f"Memproses kode ICD {processed}/{total_sirs}: {kode_icd_sirs}"
                    )

                # Filter data mapping berdasarkan kode ICD
                df_filtered = df_mapping[
                    df_mapping['kode_icd'].str.strip().str.upper() == kode_icd_sirs
                ]

                # Hitung untuk setiap pasien dengan kode ICD yang sama
                for _, mapping_row in df_filtered.iterrows():
                    gender = str(mapping_row['kelamin']).strip().upper()
                    
                    if gender not in ['L', 'P']:
                        continue

                    usia_tahun = mapping_row['usia_tahun']
                    usia_bulan = mapping_row['usia_bulan']
                    usia_hari = mapping_row['usia_hari']
                    alasan_pulang = str(mapping_row.get('alasan_pulang', '')).strip().upper()

                    # Tentukan kolom SIRS yang sesuai
                    target_col = self.get_sirs_column(usia_tahun, usia_bulan, usia_hari, gender)

                    if target_col and target_col in df_sirs.columns:
                        df_sirs.at[idx, target_col] += 1

                    # Hitung jumlah pasien keluar (hidup dan mati)
                    total_col = f"Jumlah Pasien Keluar Hidup dan Mati Menurut Jenis Kelamin_{gender}"
                    if total_col in df_sirs.columns:
                        df_sirs.at[idx, total_col] += 1

                    # Hitung jumlah pasien meninggal
                    if alasan_pulang in ['MENINGGAL', 'MATI', 'DEATH']:
                        mati_col = f"Jumlah Pasien Keluar Mati_{gender}"
                        if mati_col in df_sirs.columns:
                            df_sirs.at[idx, mati_col] += 1

                # Hitung total
                if "Jumlah Pasien Keluar Hidup dan Mati Menurut Jenis Kelamin_TOTAL" in df_sirs.columns:
                    total_l = df_sirs.at[idx, "Jumlah Pasien Keluar Hidup dan Mati Menurut Jenis Kelamin_L"]
                    total_p = df_sirs.at[idx, "Jumlah Pasien Keluar Hidup dan Mati Menurut Jenis Kelamin_P"]
                    df_sirs.at[idx, "Jumlah Pasien Keluar Hidup dan Mati Menurut Jenis Kelamin_TOTAL"] = total_l + total_p

                if "Jumlah Pasien Keluar Mati_TOTAL" in df_sirs.columns:
                    mati_l = df_sirs.at[idx, "Jumlah Pasien Keluar Mati_L"]
                    mati_p = df_sirs.at[idx, "Jumlah Pasien Keluar Mati_P"]
                    df_sirs.at[idx, "Jumlah Pasien Keluar Mati_TOTAL"] = mati_l + mati_p

            # Update preview
            self.update_progress(0.95, "Memperbarui tampilan...")
            self.df_preview = df_sirs
            self.after(0, lambda: self.show_preview(df_sirs))

            self.update_progress(1.0, "Selesai!")
            self.after(500, lambda: self.show_success("Selesai", "Optimasi SIRS berhasil!"))
            
            # Sembunyikan progress bar setelah 2 detik
            self.after(2000, self.progress_frame.pack_forget)

        except Exception as e:
            self.show_error("Error", f"Terjadi kesalahan:\n\n{e}")
        
        finally:
            self.is_processing = False
    
    def update_progress(self, value, text):
        """Update progress bar dan label"""
        self.after(0, lambda: self.progress_bar.set(value))
        self.after(0, lambda: self.progress_label.configure(text=text))
        
        # Update detail dengan persentase
        pct = int(value * 100)
        self.after(0, lambda: self.progress_detail.configure(text=f"{pct}% selesai"))
    
    def show_error(self, title, message):
        """Tampilkan error message di main thread"""
        self.after(0, lambda: messagebox.showerror(title, message))
        self.after(0, self.progress_frame.pack_forget)
    
    def show_warning(self, title, message):
        """Tampilkan warning message di main thread"""
        self.after(0, lambda: messagebox.showwarning(title, message))
        self.after(0, self.progress_frame.pack_forget)
    
    def show_success(self, title, message):
        """Tampilkan success message di main thread"""
        self.after(0, lambda: messagebox.showinfo(title, message))

    def save_result(self):
        """Simpan hasil optimasi ke file Excel"""
        if self.df_preview is None:
            messagebox.showwarning("Belum ada data", "Belum ada data untuk disimpan.")
            return

        # Buat nama file default dengan timestamp
        from datetime import datetime
        import os
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"SIRS_Hasil_{timestamp}.xlsx"
        
        # Tentukan folder default (folder excel jika ada, atau folder saat ini)
        if os.path.exists("excel"):
            default_dir = "excel"
        else:
            default_dir = os.getcwd()

        filetypes = [("Excel Files", "*.xlsx")]
        filepath = filedialog.asksaveasfilename(
            title="Simpan Hasil",
            initialfile=default_filename,
            initialdir=default_dir,
            defaultextension=".xlsx",
            filetypes=filetypes
        )

        if not filepath:
            return

        try:
            # Pastikan filepath memiliki ekstensi .xlsx
            if not filepath.endswith('.xlsx'):
                filepath += '.xlsx'
            
            # Bersihkan data dari nilai yang bermasalah
            df_clean = self.df_preview.copy()
            
            # Replace inf dan -inf dengan None
            df_clean = df_clean.replace([float('inf'), float('-inf')], None)
            
            # Simpan dengan xlsxwriter untuk lebih stabil
            try:
                with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                    df_clean.to_excel(writer, index=False, sheet_name='SIRS')
                    
                    # Format worksheet
                    workbook = writer.book
                    worksheet = writer.sheets['SIRS']
                    
                    # Header format
                    header_format = workbook.add_format({
                        'bold': True,
                        'bg_color': '#D3D3D3',
                        'border': 1
                    })
                    
                    # Auto-adjust column width
                    for idx, col in enumerate(df_clean.columns):
                        max_length = max(
                            df_clean[col].astype(str).map(len).max(),
                            len(str(col))
                        )
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.set_column(idx, idx, adjusted_width)
                        
            except ImportError:
                # Fallback ke openpyxl jika xlsxwriter tidak tersedia
                with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                    df_clean.to_excel(writer, index=False, sheet_name='SIRS')
            
            messagebox.showinfo("Berhasil", f"File berhasil disimpan ke:\n{filepath}")
            
            # Tanya apakah ingin membuka file
            if messagebox.askyesno("Buka File?", "File berhasil disimpan.\n\nApakah ingin membuka file sekarang?"):
                try:
                    os.startfile(filepath)  # Windows
                except AttributeError:
                    # Untuk Linux/Mac
                    import platform
                    if platform.system() == 'Darwin':  # Mac
                        os.system(f'open "{filepath}"')
                    else:  # Linux
                        os.system(f'xdg-open "{filepath}"')
                
        except PermissionError:
            messagebox.showerror("Error", f"File tidak dapat disimpan!\n\nKemungkinan file sedang dibuka di Excel.\nSilakan tutup file tersebut dan coba lagi.")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan file:\n\n{str(e)}\n\nCoba gunakan nama file yang berbeda.")
    
    def quick_save(self):
        """Quick save tanpa dialog - simpan langsung ke folder excel"""
        if self.df_preview is None:
            messagebox.showwarning("Belum ada data", "Belum ada data untuk disimpan.")
            return
        
        from datetime import datetime
        import os
        
        # Pastikan folder excel ada
        if not os.path.exists("excel"):
            os.makedirs("excel")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filepath = os.path.join("excel", f"SIRS_Hasil_{timestamp}.xlsx")
        
        try:
            # Bersihkan data
            df_clean = self.df_preview.copy()
            df_clean = df_clean.replace([float('inf'), float('-inf')], None)
            
            # Simpan dengan metode paling sederhana
            df_clean.to_excel(filepath, index=False, sheet_name='SIRS', engine='openpyxl')
            
            messagebox.showinfo("Berhasil", f"File tersimpan di:\n{filepath}")
            
            # Auto-open
            if messagebox.askyesno("Buka File?", "Buka file sekarang?"):
                try:
                    os.startfile(filepath)
                except:
                    messagebox.showinfo("Info", f"Silakan buka manual di:\n{filepath}")
                    
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan:\n\n{str(e)}")
    
    def export_to_pdf(self):
        """Export hasil optimasi ke PDF dengan format yang rapi"""
        if self.df_preview is None:
            messagebox.showwarning("Belum ada data", "Belum ada data untuk di-export.")
            return
        
        from datetime import datetime
        import os
        
        # Cek apakah reportlab terinstall
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch
        except ImportError:
            if messagebox.askyesno(
                "Library Tidak Ditemukan",
                "Library 'reportlab' diperlukan untuk export PDF.\n\n"
                "Install sekarang? (pip install reportlab)"
            ):
                import subprocess
                try:
                    subprocess.check_call(["pip", "install", "reportlab"])
                    messagebox.showinfo("Berhasil", "Library berhasil diinstall!\nSilakan coba export lagi.")
                except:
                    messagebox.showerror("Gagal", "Gagal install library.\nSilakan install manual:\npip install reportlab")
            return
        
        # Buat nama file default
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"SIRS_Report_{timestamp}.pdf"
        
        # Folder default
        if os.path.exists("excel"):
            default_dir = "excel"
        else:
            default_dir = os.getcwd()
        
        # Dialog save
        filetypes = [("PDF Files", "*.pdf")]
        filepath = filedialog.asksaveasfilename(
            title="Export ke PDF",
            initialfile=default_filename,
            initialdir=default_dir,
            defaultextension=".pdf",
            filetypes=filetypes
        )
        
        if not filepath:
            return
        
        try:
            # Pastikan ekstensi .pdf
            if not filepath.endswith('.pdf'):
                filepath += '.pdf'
            
            # Bersihkan data
            df_clean = self.df_preview.copy()
            df_clean = df_clean.replace([float('inf'), float('-inf')], None)
            df_clean = df_clean.fillna('')
            
            # Buat PDF dengan landscape orientation
            doc = SimpleDocTemplate(
                filepath,
                pagesize=landscape(A4),
                rightMargin=20,
                leftMargin=20,
                topMargin=30,
                bottomMargin=30
            )
            
            # Container untuk elements
            elements = []
            
            # Styles
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=16,
                textColor=colors.HexColor('#1f4788'),
                spaceAfter=20,
                alignment=1
            )
            
            subtitle_style = ParagraphStyle(
                'Subtitle',
                parent=styles['Heading2'],
                fontSize=12,
                textColor=colors.HexColor('#2e5090'),
                spaceAfter=10,
                spaceBefore=10
            )
            
            # Judul utama
            title = Paragraph("LAPORAN OPTIMASI SIRS", title_style)
            elements.append(title)
            
            # Info
            info_text = f"Tanggal Generate: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}"
            info = Paragraph(info_text, styles['Normal'])
            elements.append(info)
            
            total_text = f"Total Kode ICD: {len(df_clean)} | Total Kolom: {len(df_clean.columns)}"
            total = Paragraph(total_text, styles['Normal'])
            elements.append(total)
            elements.append(Spacer(1, 20))
            
            # Pisahkan kolom berdasarkan kategori
            basic_cols = []
            age_cols_L = []
            age_cols_P = []
            summary_cols = []
            
            for col in df_clean.columns:
                col_lower = col.lower()
                if 'kode icd' in col_lower or col_lower == 'no':
                    basic_cols.append(col)
                elif col.endswith('_L') and 'jumlah' not in col_lower:
                    age_cols_L.append(col)
                elif col.endswith('_P') and 'jumlah' not in col_lower:
                    age_cols_P.append(col)
                elif 'jumlah' in col_lower:
                    summary_cols.append(col)
            
            # TABEL 1: Data Kelompok Usia - LAKI-LAKI
            if age_cols_L:
                elements.append(Paragraph("TABEL 1: DISTRIBUSI USIA PASIEN LAKI-LAKI", subtitle_style))
                elements.append(Spacer(1, 10))
                
                # Bagi kolom usia L menjadi beberapa grup (max 12 kolom per halaman)
                chunk_size = 12
                for i in range(0, len(age_cols_L), chunk_size):
                    chunk_cols = basic_cols + age_cols_L[i:i+chunk_size]
                    df_chunk = df_clean[chunk_cols]
                    self._add_table_to_pdf(df_chunk, elements)
                    
                    if i + chunk_size < len(age_cols_L):
                        elements.append(Spacer(1, 15))
                        elements.append(Paragraph(f"<i>Lanjutan kolom usia Laki-Laki...</i>", styles['Italic']))
                        elements.append(Spacer(1, 10))
                
                elements.append(PageBreak())
            
            # TABEL 2: Data Kelompok Usia - PEREMPUAN
            if age_cols_P:
                elements.append(Paragraph("TABEL 2: DISTRIBUSI USIA PASIEN PEREMPUAN", subtitle_style))
                elements.append(Spacer(1, 10))
                
                # Bagi kolom usia P menjadi beberapa grup
                for i in range(0, len(age_cols_P), chunk_size):
                    chunk_cols = basic_cols + age_cols_P[i:i+chunk_size]
                    df_chunk = df_clean[chunk_cols]
                    self._add_table_to_pdf(df_chunk, elements)
                    
                    if i + chunk_size < len(age_cols_P):
                        elements.append(Spacer(1, 15))
                        elements.append(Paragraph(f"<i>Lanjutan kolom usia Perempuan...</i>", styles['Italic']))
                        elements.append(Spacer(1, 10))
                
                elements.append(PageBreak())
            
            # TABEL 3: Ringkasan Jumlah Pasien
            if summary_cols:
                elements.append(Paragraph("TABEL 3: RINGKASAN JUMLAH PASIEN", subtitle_style))
                elements.append(Spacer(1, 10))
                
                df_summary = df_clean[basic_cols + summary_cols]
                self._add_table_to_pdf(df_summary, elements, fontsize=7)
            
            # Build PDF
            doc.build(elements)
            
            messagebox.showinfo("Berhasil", f"PDF berhasil dibuat:\n{filepath}\n\nTotal halaman: {len(elements)//10 + 1}")
            
            # Tanya buka file
            if messagebox.askyesno("Buka PDF?", "PDF berhasil dibuat.\n\nBuka file sekarang?"):
                try:
                    os.startfile(filepath)
                except:
                    messagebox.showinfo("Info", f"Silakan buka manual:\n{filepath}")
        
        except PermissionError:
            messagebox.showerror("Error", "File sedang dibuka di aplikasi lain.\nSilakan tutup file tersebut.")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal export PDF:\n\n{str(e)}")
    
    def _create_pdf_single_table(self, df, doc, elements, styles):
        """Buat satu tabel PDF untuk data yang tidak terlalu lebar"""
        from reportlab.lib import colors
        from reportlab.platypus import Table, TableStyle, Paragraph
        
        # Limit kolom yang ditampilkan (max 20 kolom)
        max_cols = 20
        df_display = df.iloc[:, :max_cols] if len(df.columns) > max_cols else df
        
        # Prepare data untuk tabel
        data = []
        
        # Header
        headers = [str(col)[:20] for col in df_display.columns]  # Limit 20 char per header
        data.append(headers)
        
        # Rows
        for idx, row in df_display.iterrows():
            row_data = [str(val)[:20] if val != '' else '-' for val in row]
            data.append(row_data)
        
        # Buat tabel
        table = Table(data, repeatRows=1)
        
        # Style tabel
        table.setStyle(TableStyle([
            # Header
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            
            # Body
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#E7E6E6')]),
        ]))
        
        elements.append(table)
        
        # Catatan jika ada kolom yang tidak ditampilkan
        if len(df.columns) > max_cols:
            from reportlab.platypus import Spacer
            elements.append(Spacer(1, 10))
            note = Paragraph(
                f"<i>Catatan: Menampilkan {max_cols} dari {len(df.columns)} kolom. "
                f"Untuk data lengkap, gunakan export Excel.</i>",
                styles['Normal']
            )
            elements.append(note)
    
    def _create_pdf_multi_tables(self, df, doc, elements, styles):
        """Buat beberapa tabel PDF untuk data yang sangat lebar"""
        from reportlab.lib import colors
        from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, PageBreak
        
        # Pisah kolom menjadi grup
        basic_cols = ['No', 'Kode ICD']
        age_cols = [col for col in df.columns if any(x in col for x in ['thn_', 'bln_', 'hr_', 'Jam_'])]
        summary_cols = [col for col in df.columns if 'Jumlah' in col]
        
        # Tabel 1: Info Dasar + Kolom Usia (Laki-laki)
        age_cols_L = [col for col in age_cols if col.endswith('_L')]
        if age_cols_L:
            df_part1 = df[basic_cols + age_cols_L[:15]]  # Max 15 kolom usia
            elements.append(Paragraph("<b>TABEL 1: Data Usia Laki-Laki</b>", styles['Heading2']))
            elements.append(Spacer(1, 10))
            self._add_table_to_pdf(df_part1, elements)
            elements.append(PageBreak())
        
        # Tabel 2: Info Dasar + Kolom Usia (Perempuan)
        age_cols_P = [col for col in age_cols if col.endswith('_P')]
        if age_cols_P:
            df_part2 = df[basic_cols + age_cols_P[:15]]
            elements.append(Paragraph("<b>TABEL 2: Data Usia Perempuan</b>", styles['Heading2']))
            elements.append(Spacer(1, 10))
            self._add_table_to_pdf(df_part2, elements)
            elements.append(PageBreak())
        
        # Tabel 3: Summary (Jumlah Pasien)
        if summary_cols:
            df_part3 = df[basic_cols + summary_cols]
            elements.append(Paragraph("<b>TABEL 3: Ringkasan Jumlah Pasien</b>", styles['Heading2']))
            elements.append(Spacer(1, 10))
            self._add_table_to_pdf(df_part3, elements)
    
    def _add_table_to_pdf(self, df, elements, fontsize=6):
        """Helper untuk menambahkan tabel ke PDF"""
        from reportlab.lib import colors
        from reportlab.platypus import Table, TableStyle
        
        # Prepare data
        data = []
        
        # Headers - potong jika terlalu panjang
        headers = []
        for col in df.columns:
            col_str = str(col)
            # Singkat nama kolom yang panjang
            if len(col_str) > 15:
                col_str = col_str.replace('Jumlah Pasien Keluar', 'JPK')
                col_str = col_str.replace('Hidup dan Mati', 'H&M')
                col_str = col_str.replace('Menurut Jenis Kelamin', 'JK')
            headers.append(col_str[:20])
        data.append(headers)
        
        # Rows
        for idx, row in df.iterrows():
            row_data = []
            for val in row:
                val_str = str(val) if val != '' else '-'
                # Tampilkan 0 sebagai '-' untuk lebih clean
                if val_str == '0' or val_str == '0.0':
                    val_str = '-'
                row_data.append(val_str[:15])
            data.append(row_data)
        
        # Hitung lebar kolom dinamis
        num_cols = len(df.columns)
        available_width = 750  # Lebar landscape A4 dalam points
        col_width = available_width / num_cols
        col_widths = [col_width] * num_cols
        
        # Buat tabel dengan column widths
        table = Table(data, colWidths=col_widths, repeatRows=1)
        
        # Style
        table.setStyle(TableStyle([
            # Header styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), fontsize + 1),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('TOPPADDING', (0, 0), (-1, 0), 8),
            
            # Body styling
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), fontsize),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#E7E6E6')]),
            ('TOPPADDING', (0, 1), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
        ]))
        
        elements.append(table)