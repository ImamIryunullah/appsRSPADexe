import customtkinter as ctk
from tkinter import messagebox, filedialog, ttk
import pandas as pd
import sqlite3
import threading


class PuskesadScreen(ctk.CTkFrame):
    def __init__(self, master, db_path="database/mapping.db"):
        super().__init__(master)

        self.db_path = db_path
        self.file_path = None
        self.df_preview = None
        self.current_font_size = 11
        self.is_processing = False

        # === TITLE ===
        title = ctk.CTkLabel(self, text="Optimasi PUSKESAD", font=("Arial", 20))
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
        self.progress_frame.pack_forget()
        
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

    def upload_excel(self):
        """Upload file Excel PUSKESAD dengan multi-level headers"""
        filetypes = [("Excel Files", "*.xlsx"), ("Excel Files", "*.xls")]
        filepath = filedialog.askopenfilename(title="Pilih File Excel", filetypes=filetypes)

        if not filepath:
            return

        self.file_path = filepath
        self.label_file.configure(text=f"File dipilih: {filepath}")

        try:
            # Baca Excel dengan multi-level header (3 baris header)
            # Baris 0-2 adalah header, data mulai dari baris 3
            df = pd.read_excel(filepath, header=[0, 1, 2])
            
            # Flatten kolom multi-level menjadi single level
            # Gabungkan nama kolom yang tidak kosong
            new_columns = []
            for col in df.columns:
                # col adalah tuple (level0, level1, level2)
                col_parts = [str(c).strip() for c in col if not str(c).startswith('Unnamed')]
                col_name = ' '.join(col_parts) if col_parts else str(col[0])
                new_columns.append(col_name)
            
            df.columns = new_columns
            
            self.df_preview = df
            self.show_preview(df)
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca file Excel!\n\n{e}\n\nPastikan file memiliki format header yang benar.")

    def show_preview(self, df: pd.DataFrame):
        """Tampilkan preview data dalam Treeview"""
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
            self.table.column(col, width=120, anchor="w")

        for _, row in df.head(300).iterrows():
            self.table.insert("", "end", values=list(row))

    def filter_preview(self, event=None):
        """Filter preview berdasarkan keyword pencarian"""
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
        """Zoom in/out pada tabel"""
        self.current_font_size += delta
        self.current_font_size = max(self.current_font_size, 8)

        style = ttk.Style()
        style.configure("Treeview", font=("Arial", self.current_font_size))
        style.configure("Treeview.Heading", font=("Arial", self.current_font_size, "bold"))

    def get_puskesad_column(self, usia_tahun, usia_bulan, usia_hari, golongan):
        """Menentukan kolom umur PUSKESAD berdasarkan usia"""
        # Konversi ke integer
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

        # Konversi total ke hari untuk perhitungan
        total_hari = (usia_tahun * 365) + (usia_bulan * 30) + usia_hari

        # LOGIKA KELOMPOK UMUR PUSKESAD
        if total_hari <= 28:
            return "28 HARI"
        elif total_hari < 365:  # < 1 tahun
            return "28 HR < 1 THN"
        elif 1 <= usia_tahun <= 4:
            return "1 - 4 THN"
        elif 5 <= usia_tahun <= 14:
            return "5 - 14 THN"
        elif 15 <= usia_tahun <= 25:
            return "15 - 25 THN"
        elif 25 < usia_tahun <= 44:
            return "25 - 44 THN"
        elif 45 <= usia_tahun <= 64:
            return "45 - 64 THN"
        elif usia_tahun > 64:
            return ">64 THN"
        
        return None

    def get_golongan_column(self, golongan, sub_golongan=None):
        """Menentukan kolom golongan/status pasien"""
        golongan = str(golongan).upper().strip()
        
        # Mapping golongan
        if 'TNI AD' in golongan or golongan == 'AD':
            if sub_golongan and 'PNS' in str(sub_golongan).upper():
                return "TNI AD PNS AD"
            elif sub_golongan and 'KEL' in str(sub_golongan).upper():
                return "TNI AD KEL AD"
            else:
                return "TNI AD AD"
        
        elif 'AU' in golongan or 'AL' in golongan:
            if sub_golongan and 'PNS' in str(sub_golongan).upper():
                return "ANGKATAN LAIN PNS ( AL, AD MABES & KEMHAN)"
            elif sub_golongan and 'KEL' in str(sub_golongan).upper():
                return "ANGKATAN LAIN KEL  ( AL, AD MABES & KEMHAN)"
            else:
                return "ANGKATAN LAIN AU / AL"
        
        elif 'PURNAWIRAWAN' in golongan or 'BPJS' in golongan:
            return "PURNAWIRAWAN / BPJS UMUM (MANDIRI, PPPK, PEGAWAI SWASTA, PBI)"
        
        elif 'UMUM' in golongan:
            return "UMUM"
        
        return None

    def run_process(self):
        """Proses optimasi PUSKESAD"""
        if self.df_preview is None:
            messagebox.showwarning("Belum ada file", "Silakan upload file Excel terlebih dahulu.")
            return
        
        if self.is_processing:
            messagebox.showwarning("Sedang Proses", "Optimasi sedang berjalan, mohon tunggu...")
            return
        
        thread = threading.Thread(target=self._run_process_thread, daemon=True)
        thread.start()
    
    def _run_process_thread(self):
        """Thread worker untuk proses optimasi"""
        self.is_processing = True
        
        self.progress_frame.pack(pady=5, fill="x", padx=10)
        self.update_progress(0, "Memulai proses optimasi...")

        try:
            # Koneksi ke database
            self.update_progress(0.05, "Menghubungkan ke database...")
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Cek tabel mapping
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='mapping'")
            table_exists = cursor.fetchone()
            
            if not table_exists:
                conn.close()
                self.show_error(
                    "Database Error", 
                    "Tabel 'mapping' tidak ditemukan di database!\n\n"
                    f"Path database: {self.db_path}"
                )
                return
            
            # Ambil kolom dari tabel
            self.update_progress(0.1, "Memeriksa struktur tabel...")
            cursor.execute("PRAGMA table_info(mapping)")
            columns_info = cursor.fetchall()
            column_names = [col[1] for col in columns_info]
            
            # Cari nama kolom
            def find_column(possible_names):
                for possible in possible_names:
                    for col in column_names:
                        if col.lower().replace(" ", "").replace("_", "") == possible.lower().replace(" ", "").replace("_", ""):
                            return col
                return None
            
            col_kode_icd = find_column(["kode_icd", "KODE ICD", "kode icd", "KODE_ICD", "NO DAFTAR TERINCI", "no daftar terinci"])
            col_kelamin = find_column(["kelamin", "KELAMIN", "jenis_kelamin", "JENIS KELAMIN", "sex", "SEX"])
            col_usia_tahun = find_column(["usia_tahun", "USIA TAHUN", "usia_thn", "USIA_TAHUN"])
            col_usia_bulan = find_column(["usia_bulan", "USIA BULAN", "usia_bln", "USIA_BULAN"])
            col_usia_hari = find_column(["usia_hari", "USIA HARI", "usia_hr", "USIA_HARI"])
            col_golongan = find_column(["golongan", "GOLONGAN", "status", "STATUS", "golongan_pasien"])
            col_alasan_pulang = find_column(["alasan_pulang", "ALASAN PULANG", "status_pulang", "STATUS PULANG"])
            
            # Validasi kolom
            missing_cols = []
            if not col_kode_icd:
                missing_cols.append("KODE ICD / NO DAFTAR TERINCI")
            if not col_kelamin:
                missing_cols.append("KELAMIN/SEX")
            if not col_usia_tahun:
                missing_cols.append("USIA TAHUN")
            
            if missing_cols:
                conn.close()
                self.show_error(
                    "Kolom Tidak Ditemukan",
                    f"Kolom berikut tidak ditemukan:\n{', '.join(missing_cols)}\n\n"
                    f"Kolom tersedia:\n{', '.join(column_names)}"
                )
                return
            
            # Baca data dari database
            self.update_progress(0.15, "Membaca data dari database...")
            
            # Build query dengan kolom yang ditemukan
            cols_to_select = [f'"{col_kode_icd}"', f'"{col_kelamin}"', f'"{col_usia_tahun}"']
            col_names = ['kode_icd', 'kelamin', 'usia_tahun']
            
            if col_usia_bulan:
                cols_to_select.append(f'"{col_usia_bulan}"')
                col_names.append('usia_bulan')
            if col_usia_hari:
                cols_to_select.append(f'"{col_usia_hari}"')
                col_names.append('usia_hari')
            if col_golongan:
                cols_to_select.append(f'"{col_golongan}"')
                col_names.append('golongan')
            if col_alasan_pulang:
                cols_to_select.append(f'"{col_alasan_pulang}"')
                col_names.append('alasan_pulang')
            
            query = f"""
            SELECT {', '.join(cols_to_select)}
            FROM mapping
            WHERE "{col_kode_icd}" IS NOT NULL AND "{col_kode_icd}" != ''
            """
            
            df_mapping = pd.read_sql_query(query, conn)
            conn.close()
            
            df_mapping.columns = col_names
            
            # Tambahkan kolom default jika tidak ada
            if 'usia_bulan' not in df_mapping.columns:
                df_mapping['usia_bulan'] = 0
            if 'usia_hari' not in df_mapping.columns:
                df_mapping['usia_hari'] = 0
            if 'golongan' not in df_mapping.columns:
                df_mapping['golongan'] = 'UMUM'
            if 'alasan_pulang' not in df_mapping.columns:
                df_mapping['alasan_pulang'] = ''

            if df_mapping.empty:
                self.show_warning("Data Kosong", "Tidak ada data di tabel mapping!")
                return

            total_rows_db = len(df_mapping)
            self.update_progress(0.2, f"Data mapping dimuat: {total_rows_db} baris")

            # Copy dataframe PUSKESAD
            df_puskesad = self.df_preview.copy()

            # Cari kolom kode ICD di PUSKESAD
            icd_col = None
            for col in df_puskesad.columns:
                col_lower = col.lower().strip()
                if 'no daftar terinci' in col_lower or 'kode icd' in col_lower or col_lower == 'kode':
                    icd_col = col
                    break
            
            if icd_col is None:
                self.show_error("Error", "Kolom 'NO DAFTAR TERINCI' atau 'KODE ICD' tidak ditemukan!")
                return

            # Definisi kolom PUSKESAD berdasarkan screenshot
            UMUR_COLUMNS = [
                "28 HARI",
                "28 HR < 1 THN",
                "1 - 4 THN",
                "5 - 14 THN",
                "15 - 25 THN",
                "25 - 44 THN",
                "45 - 64 THN",
                ">64 THN"
            ]
            
            GOLONGAN_COLUMNS = [
                "TNI AD AD",
                "TNI AD PNS AD",
                "TNI AD KEL AD",
                "ANGKATAN LAIN AU / AL",
                "ANGKATAN LAIN PNS ( AL, AD MABES & KEMHAN)",
                "ANGKATAN LAIN KEL  ( AL, AD MABES & KEMHAN)",
                "PURNAWIRAWAN / BPJS UMUM (MANDIRI, PPPK, PEGAWAI SWASTA, PBI)",
                "UMUM"
            ]
            
            SEX_COLUMNS = ["LK", "PR"]

            self.update_progress(0.25, "Menyiapkan kolom PUSKESAD...")
            
            # Reset semua kolom numerik ke 0
            for col in UMUR_COLUMNS + GOLONGAN_COLUMNS + SEX_COLUMNS + ["JUMLAH PASIEN KELUAR (LK+PR)", "JUMLAH PASIEN KELUAR MATI"]:
                if col not in df_puskesad.columns:
                    df_puskesad[col] = 0
                else:
                    df_puskesad[col] = 0

            # Proses setiap kode ICD
            total_puskesad = len(df_puskesad)
            processed = 0
            
            self.update_progress(0.3, f"Memproses {total_puskesad} kode ICD...")
            
            for idx, puskesad_row in df_puskesad.iterrows():
                processed += 1
                progress_pct = 0.3 + (0.6 * processed / total_puskesad)
                
                kode_icd = str(puskesad_row[icd_col]).strip().upper()
                
                if kode_icd == "" or kode_icd == "NAN":
                    continue

                if processed % 10 == 0 or processed == total_puskesad:
                    self.update_progress(
                        progress_pct, 
                        f"Memproses {processed}/{total_puskesad}: {kode_icd}"
                    )

                # Filter data mapping
                df_filtered = df_mapping[
                    df_mapping['kode_icd'].str.strip().str.upper() == kode_icd
                ]

                # Hitung untuk setiap pasien
                for _, mapping_row in df_filtered.iterrows():
                    kelamin = str(mapping_row['kelamin']).strip().upper()
                    
                    # Mapping jenis kelamin ke LK/PR
                    if kelamin in ['L', 'LAKI-LAKI', 'M', 'MALE']:
                        sex_col = "LK"
                    elif kelamin in ['P', 'PEREMPUAN', 'F', 'FEMALE']:
                        sex_col = "PR"
                    else:
                        continue

                    # Hitung kolom umur
                    umur_col = self.get_puskesad_column(
                        mapping_row['usia_tahun'],
                        mapping_row['usia_bulan'],
                        mapping_row['usia_hari'],
                        mapping_row.get('golongan', '')
                    )

                    if umur_col and umur_col in df_puskesad.columns:
                        df_puskesad.at[idx, umur_col] += 1

                    # Hitung kolom golongan
                    golongan_col = self.get_golongan_column(
                        mapping_row.get('golongan', 'UMUM'),
                        None
                    )

                    if golongan_col and golongan_col in df_puskesad.columns:
                        df_puskesad.at[idx, golongan_col] += 1

                    # Hitung SEX (LK/PR)
                    if sex_col in df_puskesad.columns:
                        df_puskesad.at[idx, sex_col] += 1

                    # Hitung jumlah total
                    if "JUMLAH PASIEN KELUAR (LK+PR)" in df_puskesad.columns:
                        df_puskesad.at[idx, "JUMLAH PASIEN KELUAR (LK+PR)"] += 1

                    # Hitung pasien meninggal
                    alasan_pulang = str(mapping_row.get('alasan_pulang', '')).strip().upper()
                    if alasan_pulang in ['MENINGGAL', 'MATI', 'DEATH']:
                        if "JUMLAH PASIEN KELUAR MATI" in df_puskesad.columns:
                            df_puskesad.at[idx, "JUMLAH PASIEN KELUAR MATI"] += 1

            # Update preview
            self.update_progress(0.95, "Memperbarui tampilan...")
            self.df_preview = df_puskesad
            self.after(0, lambda: self.show_preview(df_puskesad))

            self.update_progress(1.0, "Selesai!")
            self.after(500, lambda: self.show_success("Selesai", "Optimasi PUSKESAD berhasil!"))
            
            self.after(2000, self.progress_frame.pack_forget)

        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            self.show_error("Error", f"Terjadi kesalahan:\n\n{e}\n\nDetail:\n{error_detail}")
        
        finally:
            self.is_processing = False
    
    def update_progress(self, value, text):
        """Update progress bar"""
        self.after(0, lambda: self.progress_bar.set(value))
        self.after(0, lambda: self.progress_label.configure(text=text))
        pct = int(value * 100)
        self.after(0, lambda: self.progress_detail.configure(text=f"{pct}% selesai"))
    
    def show_error(self, title, message):
        """Tampilkan error"""
        self.after(0, lambda: messagebox.showerror(title, message))
        self.after(0, self.progress_frame.pack_forget)
    
    def show_warning(self, title, message):
        """Tampilkan warning"""
        self.after(0, lambda: messagebox.showwarning(title, message))
        self.after(0, self.progress_frame.pack_forget)
    
    def show_success(self, title, message):
        """Tampilkan success"""
        self.after(0, lambda: messagebox.showinfo(title, message))

    def save_result(self):
        """Simpan hasil optimasi"""
        if self.df_preview is None:
            messagebox.showwarning("Belum ada data", "Belum ada data untuk disimpan.")
            return

        from datetime import datetime
        import os
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"PUSKESAD_Hasil_{timestamp}.xlsx"
        
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
            if not filepath.endswith('.xlsx'):
                filepath += '.xlsx'
            
            df_clean = self.df_preview.copy()
            df_clean = df_clean.replace([float('inf'), float('-inf')], None)
            
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                df_clean.to_excel(writer, index=False, sheet_name='PUSKESAD')
            
            messagebox.showinfo("Berhasil", f"File berhasil disimpan:\n{filepath}")
            
            if messagebox.askyesno("Buka File?", "Buka file sekarang?"):
                try:
                    os.startfile(filepath)
                except:
                    pass
                
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan:\n\n{str(e)}")
    
    def quick_save(self):
        """Quick save ke folder excel"""
        if self.df_preview is None:
            messagebox.showwarning("Belum ada data", "Belum ada data untuk disimpan.")
            return
        
        from datetime import datetime
        import os
        
        if not os.path.exists("excel"):
            os.makedirs("excel")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filepath = os.path.join("excel", f"PUSKESAD_Hasil_{timestamp}.xlsx")
        
        try:
            df_clean = self.df_preview.copy()
            df_clean = df_clean.replace([float('inf'), float('-inf')], None)
            df_clean.to_excel(filepath, index=False, sheet_name='PUSKESAD', engine='openpyxl')
            
            messagebox.showinfo("Berhasil", f"File tersimpan:\n{filepath}")
            
            if messagebox.askyesno("Buka File?", "Buka file sekarang?"):
                try:
                    os.startfile(filepath)
                except:
                    pass
                    
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan:\n\n{str(e)}")