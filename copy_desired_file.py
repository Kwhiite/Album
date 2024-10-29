from pathlib import Path
import shutil
import tkinter as tk
from tkinter import messagebox

class CopyDesiredFile:
    def __init__(self, root) -> None:
        self.root = root
        self.UI_component = {}
        self.setup_UI()

    def setup_UI(self):
        # Input folder
        self.create_label_entry("Source_folder:", 0)
        source_folder_entry = self.create_entry(0)

        # Desired folder name
        self.create_label_entry("Desired Folder Name:", 1)
        desired_folder_entry = self.create_entry(1)

        # Output folder
        self.create_label_entry("Output Folder:", 2)
        output_folder_entry = self.create_entry(2)

        # Convert button
        convert_button = tk.Button(self.root, text="Convert", width=20, height=3, relief="groove")
        convert_button.grid(row=3, column=1, rowspan=2, padx=5, pady=10, sticky="e")
        
        # quit button
        quit_button = tk.Button(self.root, text="Quit", width=20, height=3, relief="groove", command= self.root.quit)
        quit_button.grid(row=3, column=2, rowspan=2, padx=5, pady=10, sticky="e")

        self.UI_component = {
            "source_folder_entry": source_folder_entry,
            "desired_folder_entry": desired_folder_entry,
            "output_folder_entry": output_folder_entry,
            "convert_button": convert_button,
            "quit_button": quit_button,
        }
        
        self.start_app()

    def create_label_entry(self, text, row):
        tk.Label(self.root, text=text).grid(row=row, column=0, padx=10, pady=5, sticky="w")

    def create_entry(self, row):
        entry = tk.Entry(self.root, width=50)
        entry.grid(row=row, column=1, columnspan=2, padx=5, pady=5)
        return entry

    def find_and_copy_folders(self, root_folder, desired_folder_name, dest_folder):
        root_folder = Path(root_folder)
        dest_folder = Path(dest_folder)

        for folder in root_folder.rglob(desired_folder_name):
            if folder.is_dir():
                for file in folder.iterdir():
                    if file.is_file():
                        dest_file = dest_folder / file.name
                        shutil.copy2(file, dest_file)

    def path_check(self, source_folder, desired_folder, output_folder):
        source_folder = Path(source_folder)
        output_folder = Path(output_folder)

        if not source_folder.is_dir():
            messagebox.showerror("Error", f"{source_folder} is not a valid path. Please enter a valid directory.")
            return False
        if not desired_folder:
            messagebox.showerror("Error", "Desired folder name is empty.")
            return False
        if not output_folder:
            messagebox.showerror("Error", "Output folder is empty.")
            return False

        output_folder.mkdir(parents=True, exist_ok=True)
        return True

    def start_app(self):
        def on_convert():
            source_folder = self.UI_component["source_folder_entry"].get().strip()
            output_folder = self.UI_component["output_folder_entry"].get().strip()
            desired_folder_name = self.UI_component["desired_folder_entry"].get().strip()

            if self.path_check(source_folder, desired_folder_name, output_folder):
                # Disable button during processing
                self.UI_component["convert_button"].config(text="Processing...", state=tk.DISABLED)
                self.find_and_copy_folders(source_folder, desired_folder_name, output_folder)
                # Re-enable button after completion
                self.UI_component["convert_button"].config(text="Convert", state=tk.NORMAL)

        self.UI_component["convert_button"].config(command=on_convert)

if __name__ == "__main__":
    outer_root = tk.Tk()
    outer_root.title("Copy Desired File")
    CopyDesiredFile(outer_root)
    outer_root.mainloop()