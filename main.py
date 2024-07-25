import tkinter as tk
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import os
import string

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Processor")
        self.label = tk.Label(root, text="Drag and Drop an Excel file here")
        self.label.pack(padx=10, pady=10)

        # Bind the drop event
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.drop)

    def drop(self, event):
        file_path = event.data.strip('{}')
        self.process_file(file_path)
    def replace_special_chars(self, value):
        if pd.isna(value):
            return "Aucun"
        if isinstance(value, str) and len(value) == 1 and value in string.punctuation:
            return "Aucun"
        return value
    def process_file(self, file_path):
        try:
            # Load the Excel file
            df = pd.read_excel(file_path)

            # Add the "TEST" column with "OK" values
            df['Centre de coût principal'] = df['Centre de coût principal'].str.upper()

            column_name = "Centre de coût secondaire / service"
            if column_name in df.columns:
                df[column_name] = df[column_name].apply(self.replace_special_chars)

            print(df)
            # Construct the new file path
            dir_name = os.path.dirname(file_path)
            base_name = os.path.basename(file_path)
            new_file_path = os.path.join(dir_name, f"clean_{base_name}")

            # Save the modified DataFrame to a new Excel file
            df.to_excel(new_file_path, index=False)

            self.label.config(text=f"File processed and saved as {new_file_path}")

        except Exception as e:
            self.label.config(text=f"Error: {e}")


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()
