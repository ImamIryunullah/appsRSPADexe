import customtkinter as ctk
from tkinter import messagebox, filedialog, ttk
import pandas as pd
import sqlite3
import threading
import re


class PuskesadScreen(ctk.CTkFrame):
    def __init__(self, master, db_path="database/mapping.db"):
        super().__init__(master)

        self.db_path = db_path
        self.file_path = None
        self.df_preview = None
        self.df_cleaned = None  # DataFrame dengan kode ICD yang sudah di-expand
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

        # === CLEAN BUTTON ===
        btn_clean = ctk.CTkButton(self, text="Clean & Expand Kode ICD", command=self.clean_icd_codes, fg_color="orange")
        btn_clean.pack(pady=5)

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
            self.df_cleaned = None  # Reset cleaned data
            self.show_preview(df)
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca file Excel!\n\n{e}\n\nPastikan file memiliki format header yang benar.")

    def expand_icd_code(self, code_str):
        """
        Expand kode ICD yang tidak standar menjadi list kode ICD standar

        Contoh:
        - "A 00" -> ["A00.0"]
        - "A 06.0-.3,.5-.9" -> ["A06.0", "A06.1", "A06.2", "A06.3", "A06.5", "A06.6", "A06.7", "A06.8", "A06.9"]
        - "A 15.1-16.2" -> ["A15.1", "A15.2", ..., "A15.9", "A16.0", "A16.1", "A16.2"]
        - "A 18.1.3-.8" -> ["A18.1.3", "A18.1.4", ..., "A18.1.8"]
        - "A 02, 04-05, A 07-08" -> ["A02.0", "A04.0", "A04.1", ..., "A04.9", "A05.0", "A07.0", "A08.0"]
        """
        if pd.isna(code_str) or str(code_str).strip() == "":
            return []

        code_str = str(code_str).strip().upper()
        expanded_codes = []
        
        # Cari prefix huruf di awal string (untuk digunakan di part tanpa huruf)
        main_prefix_match = re.match(r'^([A-Z])\s+', code_str)
        main_prefix = main_prefix_match.group(1) if main_prefix_match else None

        # Split berdasarkan koma
        parts = [p.strip() for p in code_str.split(',')]

        for part in parts:
            if '-' in part:
                # Proses range
                self._expand_range(part, expanded_codes, main_prefix)
            else:
                # Proses single code
                self._expand_single(part, expanded_codes, main_prefix)

        return expanded_codes

    def _expand_single(self, code_str, result_list, main_prefix=None):
        """Expand single code (tanpa range)"""
        clean = re.sub(r'\s+', '', code_str)
        
        # Jika tidak ada huruf tapi ada main_prefix, tambahkan prefix
        if not re.search(r'[A-Z]', clean) and main_prefix:
            clean = main_prefix + clean

        if '.' not in clean:
            # Cari pattern A00 atau A000 atau A00000
            match = re.match(r'([A-Z])(\d+)', clean)
            if match:
                letter = match.group(1)
                number = int(match.group(2))
                result_list.append(f"{letter}{number:02d}.0")
            else:
                result_list.append(clean)
        else:
            result_list.append(clean)

    def _expand_range(self, range_str, result_list, main_prefix=None):
        """
        Expand range codes dengan berbagai format:
        - "A 04-05" -> A04.0, A04.1, ..., A04.9, A05.0
        - "04-05" (dengan main_prefix A) -> A04.0, A04.1, ..., A04.9, A05.0
        - "A 06.0-.3" -> A06.0, A06.1, A06.2, A06.3
        - ".5-.9" (dengan main_prefix A dan context A06) -> A06.5, A06.6, A06.7, A06.8, A06.9
        - "A 15.1-16.2" -> A15.1, A15.2, ..., A16.2
        - "A 18.1.3-.8" -> A18.1.3, A18.1.4, ..., A18.1.8
        """
        # Cari prefix huruf di awal
        prefix_match = re.match(r'([A-Z])\s*', range_str)
        if prefix_match:
            prefix = prefix_match.group(1)
            rest = range_str[len(prefix_match.group(0)):].strip()
        elif main_prefix:
            # Gunakan main_prefix jika tidak ada prefix di range ini
            prefix = main_prefix
            rest = range_str.strip()
        else:
            result_list.append(range_str)
            return

        # Coba pattern: ".5-.9" (hanya sub-code range, butuh context dari result sebelumnya)
        match = re.match(r'^\.(\d+)-\.(\d+)$', rest)
        if match:
            start_sub = int(match.group(1))
            end_sub = int(match.group(2))
            
            # Ambil main code dari result terakhir
            if result_list:
                last_code = result_list[-1]
                last_match = re.match(r'([A-Z])(\d+)', last_code)
                if last_match:
                    main_code = int(last_match.group(2))
                    for sub in range(start_sub, end_sub + 1):
                        result_list.append(f"{prefix}{main_code:02d}.{sub}")
                    return
            
            # Fallback jika tidak ada context
            result_list.append(range_str)
            return

        # Coba pattern: "06.0-.3" (same main, range sub)
        match = re.match(r'^(\d+)\.(\d+)-\.(\d+)$', rest)
        if match:
            main_code = int(match.group(1))
            start_sub = int(match.group(2))
            end_sub = int(match.group(3))
            for sub in range(start_sub, end_sub + 1):
                result_list.append(f"{prefix}{main_code:02d}.{sub}")
            return

        # Coba pattern: "15.1-16.2" (cross main code range)
        match = re.match(r'^(\d+)\.(\d+)-(\d+)\.(\d+)$', rest)
        if match:
            start_main = int(match.group(1))
            start_sub = int(match.group(2))
            end_main = int(match.group(3))
            end_sub = int(match.group(4))

            # Generate range meliputi multiple main codes
            for main in range(start_main, end_main + 1):
                if main == start_main:
                    # Mulai dari start_sub sampai 9
                    for sub in range(start_sub, 10):
                        result_list.append(f"{prefix}{main:02d}.{sub}")
                elif main == end_main:
                    # Dari 0 sampai end_sub
                    for sub in range(0, end_sub + 1):
                        result_list.append(f"{prefix}{main:02d}.{sub}")
                else:
                    # Lengkap 0-9
                    for sub in range(0, 10):
                        result_list.append(f"{prefix}{main:02d}.{sub}")
            return

        # Coba pattern: "18.1.3-.8" (three-level range)
        match = re.match(r'^(\d+)\.(\d+)\.(\d+)-\.(\d+)$', rest)
        if match:
            main_code = int(match.group(1))
            sub_code = int(match.group(2))
            start_third = int(match.group(3))
            end_third = int(match.group(4))
            for third in range(start_third, end_third + 1):
                result_list.append(f"{prefix}{main_code:02d}.{sub_code}.{third}")
            return

        # Coba pattern: "04-05" (main code only range, expand semua sub 0-9)
        match = re.match(r'^(\d+)-(\d+)$', rest)
        if match:
            start_main = int(match.group(1))
            end_main = int(match.group(2))
            for main in range(start_main, end_main + 1):
                # Expand semua sub-code dari 0-9
                for sub in range(0, 10):
                    result_list.append(f"{prefix}{main:02d}.{sub}")
            return

        # Fallback
        clean = prefix + re.sub(r'\s+', '', rest)
        if '.' not in clean:
            match = re.match(r'([A-Z])(\d+)', clean)
            if match:
                letter = match.group(1)
                number = int(match.group(2))
                result_list.append(f"{letter}{number:02d}.0")
            else:
                result_list.append(clean)
        else:
            result_list.append(clean)

    def clean_icd_codes(self):
        """Clean dan expand kode ICD di kolom NO DAFTAR TERINCI (tetap 1 baris, multi-line dalam sel)"""
        if self.df_preview is None:
            messagebox.showwarning("Belum ada file", "Silakan upload file Excel terlebih dahulu.")
            return
        
        try:
            # Cari kolom NO DAFTAR TERINCI
            icd_col = None
            for col in self.df_preview.columns:
                col_lower = col.lower().strip()
                if 'no daftar terinci' in col_lower or 'kode icd' in col_lower or col_lower == 'kode':
                    icd_col = col
                    break
            
            if icd_col is None:
                messagebox.showerror("Error", "Kolom 'NO DAFTAR TERINCI' tidak ditemukan!")
                return
            
            # Buat DataFrame baru dengan kode ICD yang di-expand dalam satu sel
            self.df_cleaned = self.df_preview.copy()
            
            # Cari index kolom NO DAFTAR TERINCI
            cols = list(self.df_cleaned.columns)
            icd_col_idx = cols.index(icd_col)
            
            # Simpan kode asli terlebih dahulu
            kode_asli_series = self.df_cleaned[icd_col].copy()
            
            # Expand kode ICD menjadi multi-line string
            total_expanded = 0
            for idx, row in self.df_cleaned.iterrows():
                original_code = row[icd_col]
                expanded_codes = self.expand_icd_code(original_code)
                
                if expanded_codes and len(expanded_codes) > 0:
                    # Gabungkan dengan koma untuk pemisah
                    self.df_cleaned.at[idx, icd_col] = ', '.join(expanded_codes)
                    total_expanded += len(expanded_codes)
                else:
                    # Tetap gunakan kode asli jika tidak bisa expand
                    self.df_cleaned.at[idx, icd_col] = str(original_code) if pd.notna(original_code) else ''
            
            # Tambahkan kolom KODE ASLI setelah NO DAFTAR TERINCI
            cols.insert(icd_col_idx + 1, 'KODE ASLI')
            self.df_cleaned['KODE ASLI'] = kode_asli_series
            self.df_cleaned = self.df_cleaned[cols]
            
            # Update preview
            self.show_preview(self.df_cleaned)
            
            messagebox.showinfo(
                "Berhasil", 
                f"Kode ICD berhasil dibersihkan!\n\n"
                f"Total baris data: {len(self.df_cleaned)}\n"
                f"Total kode ICD hasil expand: {total_expanded}\n\n"
                f"Kolom 'NO DAFTAR TERINCI' sekarang berisi kode yang sudah di-expand (dipisah koma).\n"
                f"Kolom 'KODE ASLI' berisi kode sebelum cleaning."
            )
            
        except Exception as e:
            import traceback
            messagebox.showerror("Error", f"Gagal membersihkan kode ICD:\n\n{e}\n\n{traceback.format_exc()}")

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

        # Gunakan df_cleaned jika ada, jika tidak gunakan df_preview
        df_to_filter = self.df_cleaned if self.df_cleaned is not None else self.df_preview
        
        df_filtered = df_to_filter[
            df_to_filter.apply(lambda row: row.astype(str).str.lower().str.contains(keyword).any(), axis=1)
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

    def run_process(self):
        """Proses optimasi PUSKESAD"""
        if self.df_preview is None:
            messagebox.showwarning("Belum ada file", "Silakan upload file Excel terlebih dahulu.")
            return
        
        if self.df_cleaned is None:
            if messagebox.askyesno(
                "Belum di-clean", 
                "Kode ICD belum dibersihkan.\n\n"
                "Lanjutkan dengan data asli tanpa cleaning?\n"
                "(Disarankan klik 'Clean & Expand Kode ICD' terlebih dahulu)"
            ):
                pass
            else:
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
            # Gunakan self.df_cleaned jika tersedia, jika tidak gunakan self.df_preview
            df_to_process = self.df_cleaned if self.df_cleaned is not None else self.df_preview
            
            self.update_progress(0.1, "Membaca data dari database...")
            
            # Baca data dari database mapping
            conn = sqlite3.connect(self.db_path)
            df_mapping = pd.read_sql_query("SELECT * FROM mapping", conn)
            conn.close()
            
            self.update_progress(0.2, f"Data mapping terbaca: {len(df_mapping)} baris")
            
            # Cari kolom NO DAFTAR TERINCI di PUSKESAD
            icd_col = None
            for col in df_to_process.columns:
                col_lower = col.lower().strip()
                if 'no daftar terinci' in col_lower or 'kode icd' in col_lower:
                    icd_col = col
                    break
            
            if icd_col is None:
                self.show_error("Error", "Kolom 'NO DAFTAR TERINCI' tidak ditemukan!")
                return
            
            # Inisialisasi hasil dengan copy dataframe
            df_result = df_to_process.copy()
            
            # Mapping kolom berdasarkan ANGKATAN
            angkatan_mapping = {
                'AD': 'PASIEN MENURUT GOLONGAN / STATUS TNI AD AD',
                'PNS AD': 'PASIEN MENURUT GOLONGAN / STATUS TNI AD PNS AD',
                'KEL AD': 'PASIEN MENURUT GOLONGAN / STATUS TNI AD KEL AD',
                'AU': 'PASIEN MENURUT GOLONGAN / STATUS ANGKATAN LAIN AU / AL',
                'AL': 'PASIEN MENURUT GOLONGAN / STATUS ANGKATAN LAIN AU / AL',
                'TNI': 'PASIEN MENURUT GOLONGAN / STATUS ANGKATAN LAIN AU / AL',
                'PNS AL': 'PASIEN MENURUT GOLONGAN / STATUS ANGKATAN LAIN PNS ( AL, AD MABES & KEMHAN)',
                'KEMENTERIAN': 'PASIEN MENURUT GOLONGAN / STATUS ANGKATAN LAIN PNS ( AL, AD MABES & KEMHAN)',
                'KEL AL': 'PASIEN MENURUT GOLONGAN / STATUS ANGKATAN LAIN KEL  ( AL, AD MABES & KEMHAN)',
            }
            
            # Inisialisasi kolom numerik dengan 0
            numeric_cols = [
                'PASIEN MENURUT GOLONGAN / STATUS TNI AD AD',
                'PASIEN MENURUT GOLONGAN / STATUS TNI AD PNS AD',
                'PASIEN MENURUT GOLONGAN / STATUS TNI AD KEL AD',
                'PASIEN MENURUT GOLONGAN / STATUS ANGKATAN LAIN AU / AL',
                'PASIEN MENURUT GOLONGAN / STATUS ANGKATAN LAIN PNS ( AL, AD MABES & KEMHAN)',
                'PASIEN MENURUT GOLONGAN / STATUS ANGKATAN LAIN KEL  ( AL, AD MABES & KEMHAN)',
                'PASIEN MENURUT GOLONGAN / STATUS PURNAWIRAWAN / BPJS UMUM (MANDIRI, PPPK, PEGAWAI SWASTA, PBI)',
                'PASIEN MENURUT GOLONGAN / STATUS UMUM',
                'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 28 HARI',
                'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 28 HR < 1 THN',
                'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 1 - 4 THN',
                'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 5 - 14 THN',
                'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 15 - 25 THN',
                'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 25 - 44 THN',
                'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 45 - 64 THN',
                'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR >64 THN',
                'PASIEN KELUAR (HIDUP & MATI) MENURUT SEX LK',
                'PASIEN KELUAR (HIDUP & MATI) MENURUT SEX PR',
                'JUMLAH PASIEN KELUAR (LK+PR)',
                'JUMLAH PASIEN KELUAR MATI',
            ]
            
            for col in numeric_cols:
                if col in df_result.columns:
                    df_result[col] = 0
            
            total_rows = len(df_result)
            self.update_progress(0.3, "Memproses matching kode ICD...")
            
            # Proses setiap baris PUSKESAD
            for idx, row in df_result.iterrows():
                progress = 0.3 + (0.6 * (idx / total_rows))
                if idx % 10 == 0:
                    self.update_progress(progress, f"Memproses baris {idx+1}/{total_rows}...")
                
                # Ambil kode ICD yang sudah di-expand (dipisah koma)
                icd_codes_str = str(row[icd_col])
                if pd.isna(icd_codes_str) or icd_codes_str.strip() == '' or icd_codes_str == 'nan':
                    continue
                
                # Split kode ICD (sudah di-expand, dipisah koma)
                icd_codes = [code.strip() for code in icd_codes_str.split(',') if code.strip()]
                
                # Untuk setiap kode ICD, cari di mapping
                for icd_code in icd_codes:
                    # Cari matching di database mapping
                    matching_rows = df_mapping[df_mapping['kode_icd'].str.strip().str.upper() == icd_code.upper()]
                    
                    if len(matching_rows) == 0:
                        continue
                    
                    # Proses setiap matching record
                    for _, map_row in matching_rows.iterrows():
                        angkatan = str(map_row.get('angkatan', '')).strip().upper()
                        jenis_pembayaran = str(map_row.get('jenis_pembayaran_', '')).strip().upper()
                        
                        # Tentukan kolom target berdasarkan ANGKATAN
                        target_col = None
                        
                        if angkatan in angkatan_mapping:
                            target_col = angkatan_mapping[angkatan]
                        elif 'LAIN' in angkatan or 'PPPK' in angkatan:
                            # LAIN-LAIN dan PPPK masuk ke kolom berdasarkan jenis pembayaran
                            if 'BPJS' in jenis_pembayaran or 'DINAS' in jenis_pembayaran:
                                target_col = 'PASIEN MENURUT GOLONGAN / STATUS PURNAWIRAWAN / BPJS UMUM (MANDIRI, PPPK, PEGAWAI SWASTA, PBI)'
                            else:
                                target_col = 'PASIEN MENURUT GOLONGAN / STATUS UMUM'
                        
                        # Jika kolom target ditemukan, tambahkan count
                        if target_col and target_col in df_result.columns:
                            df_result.at[idx, target_col] += 1
                        
                        # Tambahan: cek jenis pembayaran untuk kolom BPJS/UMUM
                        if 'BPJS' in jenis_pembayaran or 'DINAS' in jenis_pembayaran:
                            bpjs_col = 'PASIEN MENURUT GOLONGAN / STATUS PURNAWIRAWAN / BPJS UMUM (MANDIRI, PPPK, PEGAWAI SWASTA, PBI)'
                            if bpjs_col in df_result.columns and target_col != bpjs_col:
                                df_result.at[idx, bpjs_col] += 1
                        elif 'UMUM' in jenis_pembayaran:
                            umum_col = 'PASIEN MENURUT GOLONGAN / STATUS UMUM'
                            if umum_col in df_result.columns and target_col != umum_col:
                                df_result.at[idx, umum_col] += 1
                        
                        # === PROSES UMUR ===
                        # Ambil data umur dari mapping
                        usia_tahun = str(map_row.get('usia_tahun', '0')).strip()
                        usia_bulan = str(map_row.get('usia_bulan', '0')).strip()
                        usia_hari = str(map_row.get('usia_hari', '0')).strip()
                        
                        # Konversi ke integer (default 0 jika kosong atau invalid)
                        try:
                            tahun = int(float(usia_tahun)) if usia_tahun and usia_tahun != '' else 0
                        except:
                            tahun = 0
                        
                        try:
                            bulan = int(float(usia_bulan)) if usia_bulan and usia_bulan != '' else 0
                        except:
                            bulan = 0
                        
                        try:
                            hari = int(float(usia_hari)) if usia_hari and usia_hari != '' else 0
                        except:
                            hari = 0
                        
                        # Tentukan kolom umur berdasarkan usia
                        # Prioritas: tahun > bulan > hari
                        umur_col = None
                        
                        if tahun == 0 and bulan == 0 and hari > 0 and hari <= 28:
                            # 0-28 hari
                            umur_col = 'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 28 HARI'
                        elif tahun == 0 and ((bulan > 0 and bulan < 12) or (bulan == 0 and hari > 28)):
                            # 28 hari < umur < 1 tahun (bulan 1-11 atau hari > 28)
                            umur_col = 'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 28 HR < 1 THN'
                        elif tahun >= 1 and tahun <= 4:
                            # 1-4 tahun (termasuk 1 tahun sampai 4 tahun 11 bulan 30 hari)
                            umur_col = 'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 1 - 4 THN'
                        elif tahun >= 5 and tahun <= 14:
                            # 5-14 tahun
                            umur_col = 'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 5 - 14 THN'
                        elif tahun >= 15 and tahun <= 25:
                            # 15-25 tahun
                            umur_col = 'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 15 - 25 THN'
                        elif tahun > 25 and tahun <= 44:
                            # >25 - 44 tahun (26-44 tahun)
                            umur_col = 'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 25 - 44 THN'
                        elif tahun >= 45 and tahun <= 64:
                            # 45-64 tahun
                            umur_col = 'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR 45 - 64 THN'
                        elif tahun > 64:
                            # >64 tahun
                            umur_col = 'PASIEN KELUAR ( HIDUP & MATI ) MENURUT GOL.UMUR >64 THN'
                        
                        # Tambahkan count ke kolom umur
                        if umur_col and umur_col in df_result.columns:
                            df_result.at[idx, umur_col] += 1
                        
                        # === PROSES JENIS KELAMIN ===
                        kelamin = str(map_row.get('kelamin', '')).strip().upper()
                        
                        if kelamin in ['L', 'LK', 'LAKI', 'LAKI-LAKI', 'M', 'MALE']:
                            sex_col = 'PASIEN KELUAR (HIDUP & MATI) MENURUT SEX LK'
                            if sex_col in df_result.columns:
                                df_result.at[idx, sex_col] += 1
                        elif kelamin in ['P', 'PR', 'PEREMPUAN', 'F', 'FEMALE', 'WANITA']:
                            sex_col = 'PASIEN KELUAR (HIDUP & MATI) MENURUT SEX PR'
                            if sex_col in df_result.columns:
                                df_result.at[idx, sex_col] += 1
                        
                        # === PROSES ALASAN PULANG (MENINGGAL) ===
                        alasan_pulang = str(map_row.get('alasan_pulang', '')).strip().upper()
                        
                        if 'MENINGGAL' in alasan_pulang or alasan_pulang in ['MATI', 'MENINGGAL DUNIA', 'DEATH', 'DIED']:
                            mati_col = 'JUMLAH PASIEN KELUAR MATI'
                            if mati_col in df_result.columns:
                                df_result.at[idx, mati_col] += 1
            
            self.update_progress(0.95, "Menghitung total dan finalisasi...")
            
            # === HITUNG JUMLAH PASIEN KELUAR (LK+PR) ===
            # Jumlahkan kolom SEX LK dan SEX PR
            lk_col = 'PASIEN KELUAR (HIDUP & MATI) MENURUT SEX LK'
            pr_col = 'PASIEN KELUAR (HIDUP & MATI) MENURUT SEX PR'
            total_col = 'JUMLAH PASIEN KELUAR (LK+PR)'
            
            if lk_col in df_result.columns and pr_col in df_result.columns and total_col in df_result.columns:
                df_result[total_col] = df_result[lk_col] + df_result[pr_col]
            
            
            # Simpan hasil ke df_cleaned
            self.df_cleaned = df_result
            
            # Update preview
            self.after(0, lambda: self.show_preview(df_result))
            
            self.update_progress(1.0, "Selesai!")
            
            # Hitung statistik
            total_matches = df_result[numeric_cols].sum().sum()
            self.after(500, lambda: self.show_success(
                "Optimasi Selesai", 
                f"Proses optimasi berhasil!\n\n"
                f"Total baris PUSKESAD: {len(df_result)}\n"
                f"Total data mapping: {len(df_mapping)}\n"
                f"Total matching: {int(total_matches)}\n\n"
                f"Silakan cek preview dan simpan hasil."
            ))
            
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
        # Prioritaskan df_cleaned, jika tidak ada gunakan df_preview
        df_to_save = self.df_cleaned if self.df_cleaned is not None else self.df_preview
        
        if df_to_save is None:
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
            
            df_clean = df_to_save.copy()
            if 'KODE ASLI' in df_clean.columns:
                df_clean = df_clean.drop(columns=['KODE ASLI'])
            
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
        # Prioritaskan df_cleaned, jika tidak ada gunakan df_preview
        df_to_save = self.df_cleaned if self.df_cleaned is not None else self.df_preview
        
        if df_to_save is None:
            messagebox.showwarning("Belum ada data", "Belum ada data untuk disimpan.")
            return
        
        from datetime import datetime
        import os
        
        if not os.path.exists("excel"):
            os.makedirs("excel")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filepath = os.path.join("excel", f"PUSKESAD_Hasil_{timestamp}.xlsx")
        
        try:
            df_clean = df_to_save.copy()
            if 'KODE ASLI' in df_clean.columns:
                df_clean = df_clean.drop(columns=['KODE ASLI'])
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