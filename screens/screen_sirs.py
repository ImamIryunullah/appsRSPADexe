import customtkinter as ctk
from tkinter import messagebox

class SirsScreen(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)

        self.title("Optimasi SIRS")
        self.geometry("500x300")

        label = ctk.CTkLabel(self, text="Optimasi SIRS", font=("Arial", 20))
        label.pack(pady=20)

        btn_run = ctk.CTkButton(self, text="Mulai Optimasi", command=self.run_process)
        btn_run.pack(pady=20)

    def run_process(self):
        # TODO: logika SIRS
        messagebox.showinfo("Selesai", "Optimasi SIRS selesai! (logika belum dibuat)")
