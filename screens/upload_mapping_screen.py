import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from logic import save_mapping_to_db
import threading

class UploadMappingScreen(tk.Frame):
    def __init__(self, master):
        super().__init__(master)

        self.label_title = ttk.Label(self, text="Upload Mapping Otomatis", font=("Arial", 16))
        self.label_title.pack(pady=20)

        self.btn_upload = ttk.Button(self, text="Pilih File Excel", command=self.open_file)
        self.btn_upload.pack(pady=10)

        self.label_file = ttk.Label(self, text="Belum ada file dipilih", foreground="gray")
        self.label_file.pack()

        # ================= PROGRESS BAR =================
        self.progress_label = ttk.Label(self, text="")
        self.progress_label.pack(pady=5)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="indeterminate", length=300)
        self.progress.pack(pady=10)
        self.progress.pack_forget()   # disembunyikan dulu

    # ======================================================
    # BUKA FILE
    # ======================================================
    def open_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        if not file_path:
            return

        self.label_file.config(text=f"File: {file_path}")

        # Jalankan upload di thread agar UI tidak freeze
        t = threading.Thread(target=self.start_upload, args=(file_path,))
        t.start()

    # ======================================================
    # MULAI UPLOAD (THREAD)
    # ======================================================
    def start_upload(self, file_path):
        # Munculkan progress bar
        self.progress_label.config(text="Uploading, mohon tunggu...")
        self.progress.pack()
        self.progress.start(10)  # animasi loading

        # Disable tombol upload selama proses
        self.btn_upload.config(state="disabled")

        try:
            total = save_mapping_to_db(file_path)

            # Setelah selesai
            self.progress.stop()
            self.progress.pack_forget()
            self.progress_label.config(text="")

            self.btn_upload.config(state="normal")

            messagebox.showinfo("Sukses", f"Berhasil upload {total} baris mapping!")

        except Exception as e:
            self.progress.stop()
            self.progress.pack_forget()
            self.progress_label.config(text="")
            self.btn_upload.config(state="normal")

            messagebox.showerror("Error", str(e))
