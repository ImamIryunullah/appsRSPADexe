import customtkinter as ctk
from tkinter import filedialog, messagebox

class UploadScreen(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        
        self.title("Upload Data Mentah (Mapping)")
        self.geometry("500x300")

        label = ctk.CTkLabel(self, text="Upload Data Mentah untuk Mapping", font=("Arial", 20))
        label.pack(pady=20)

        self.entry_file = ctk.CTkEntry(self, width=300)
        self.entry_file.pack()

        btn_browse = ctk.CTkButton(self, text="Pilih File", command=self.choose_file)
        btn_browse.pack(pady=10)

        btn_upload = ctk.CTkButton(self, text="Upload Data", command=self.upload_data)
        btn_upload.pack(pady=10)

    def choose_file(self):
        file = filedialog.askopenfilename(
            title="Pilih file data mentah",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file:
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, file)

    def upload_data(self):
        file_path = self.entry_file.get()
        if not file_path:
            messagebox.showwarning("Error", "Pilih file terlebih dahulu!")
            return

        # TODO: logika upload mapping
        messagebox.showinfo("Sukses", "Upload berhasil! (logika belum diimplementasi)")
