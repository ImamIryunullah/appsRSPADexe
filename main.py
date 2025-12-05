import tkinter as tk
from tkinter import ttk
from screens.upload_mapping_screen import UploadMappingScreen
from screens.view_mapping_screen import ViewMappingScreen

def main():
    root = tk.Tk()
    root.title("RSPAD Optimizer")
    root.geometry("900x500")

    notebook = ttk.Notebook(root)
    notebook.pack(expand=True, fill="both")

    # TAB 1: Upload Mapping
    tab_mapping = UploadMappingScreen(notebook)
    notebook.add(tab_mapping, text="Upload Mapping")

    # TAB 2: Lihat Data Mapping
    tab_view = ViewMappingScreen(notebook)
    notebook.add(tab_view, text="Lihat Data Mapping")

    root.mainloop()

if __name__ == "__main__":
    main()
