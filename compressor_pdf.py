import shutil
import subprocess
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from pathlib import Path
from PyPDF2 import PdfReader

class PDFCompressor:
    def __init__(self, root):
        self.root = root
        self.UI_component = {}
        self.stop_compression= False
        self.timeout = 300
        self.setup_UI(root)

    def stop_compression_process(self):
        self.stop_compression= True
    
    #Main PDF compression method with a timeout
    def compress_pdf(self, input_pdf_path, output_pdf_path, quality="/ebook"):
        """
        Compress the PDF file using Ghostscript, with a timeout.
        
        :param input_pdf_path: Path to the input PDF file
        :param output_pdf_path: Path to save the compressed PDF file
        :param quality: Compression quality setting. Options: /screen, /ebook, /printer, /prepress
        :param timeout: Maximum time allowed for the compression process (in seconds).
        """
        command = [
            'gs',
            '-sDEVICE=pdfwrite',
            '-dCompatibilityLevel=1.5',
            f'-dPDFSETTINGS={quality}',
            '-dNOPAUSE',
            '-dQUIET',
            '-dBATCH',
            '-dSAFER',
            f'-sOutputFile={output_pdf_path}',
            input_pdf_path
        ]
        
        try:
            # Use Popen and set a timeout to avoid excessive execution time, without using run.
            # with contextlib.ExitStack() as stack:
            #     process = stack.enter_context(subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE))
            #     stdout, stderr= process.communicate(timeout= timeout)
            #     if process.returncode != 0:
            #         raise subprocess.CalledProcessError(process.returncode, command, output=stdout, stderr=stderr)
            # 使用 Popen，設定timeout以避免過長執行，不使用run
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            process.communicate(timeout=self.timeout)
            return self.check_pdf(output_pdf_path)
        except subprocess.TimeoutExpired:
            process.kill()
            print(f"Compression for {Path(input_pdf_path).name} timed out after {self.timeout} seconds.")
            return False  
        except subprocess.CalledProcessError as e:
            print(f"Error compressing {Path(input_pdf_path).name}: {e}")
            return False

    #UI Setup
    def setup_UI(self, root):
        # Input folder
        tk.Label(root, text="Input Folder:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        input_folder_entry = tk.Entry(root, width=50)
        input_folder_entry.grid(row=0, column=1, columnspan= 2,padx=5, pady=5)
        
        # Output folder
        tk.Label(root, text="Output Folder (optional):").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        output_folder_entry = tk.Entry(root, width=50)
        output_folder_entry.grid(row=1, column=1, columnspan= 2,padx=5, pady=5)
        
        #timeout option menu
        tk.Label(root, text="Set compression time out:").grid(row=2, column=0, padx=10, pady=1, sticky="w")
        option_list = [60, 120, 180, 240, 300, 360]
        option_value = tk.IntVar()
        option_value.set(self.timeout)
        def update_timeout(value):
            self.timeout = int(value)
        option_menu = tk.OptionMenu(root, option_value, *option_list, command= update_timeout)
        option_menu.config(width=15)
        option_menu.grid(row=3, column=0, padx=10, pady=1, sticky= "we")
        
        # Convert button
        convert_button = tk.Button(root, text= "Start Compression", width= 20, height= 3, relief= "groove")
        convert_button.grid(row=2, column=1, rowspan=2, padx=5, pady=10, sticky= "e")
        
        # Quit button
        quit_button = tk.Button(root, text= "Quit the compressor", width= 20, height= 3, relief= "groove", command= root.quit)
        quit_button.grid(row= 2, column= 2, rowspan=2, padx=5, pady=10)
        
        # progress num
        progress_state_label = tk.Label(root, text= "Processed/Total files:   0/0")
        progress_state_label.grid(row=4, column=2, padx=10, pady=5, sticky= "e")
        
        # finished file check
        finish_check_label = tk.Label(root, text= "Completed file count:   Waiting...")
        finish_check_label.grid(row=5, column=2, padx=10, pady=5, sticky= "e")
        
        # progress bar
        progress_bar_label = tk.Label(root, text= "0.0 %")
        progress_bar_label.grid(row=4, column=0, columnspan=2, padx=10, pady=5)
        
        progress_bar =ttk.Progressbar(root, orient= "horizontal", length= 400, mode= "determinate")
        progress_bar.grid(row=5, column=0, columnspan=2, padx=5, pady=10)
        
        # Progress text
        tk.Label(root, text="Progress:").grid(row=6, column=0, padx=10, pady=5, sticky="w")
        progress_text = tk.Text(root, height=10, width=30)
        progress_text.grid(row=7, column=0, padx=10, pady=0)
        
        #Broken text
        tk.Label(root, text="Broken file:").grid(row=6, column=1, padx=10, pady=5, sticky="w")
        broken_text = tk.Text(root, height=10, width=30)
        broken_text.grid(row=7, column=1, padx=10, pady=0)
        
        #Not PDF text
        tk.Label(root, text="Not PDF:").grid(row=6, column=2, padx=10, pady=5, sticky="w")
        not_PDF_text = tk.Text(root, height=10, width=30)
        not_PDF_text.grid(row=7, column=2, padx=10, pady=0)

        self.UI_component= {
            "input_folder_entry": input_folder_entry,
            "output_folder_entry": output_folder_entry,
            "convert_button": convert_button,
            "quit_button": quit_button,
            "progress_state_label": progress_state_label,
            "finish_check_label": finish_check_label,
            "progress_bar_label": progress_bar_label,
            "progress_bar": progress_bar,
            "progress_text": progress_text,
            "broken_text": broken_text,
            "not_PDF_text": not_PDF_text
        }
        self.start_compression()

    #Update to UI_component
    def update_UI_component(self, component_name, content= None, is_process_bar= False):
        try:
            component= self.UI_component[component_name]
            if is_process_bar:
                component.step(1)
            else:
                if component_name in ["progress_state_label", "finish_check_label", "progress_bar_label"]:
                    component.config(text= content)
                else:
                    component.insert(tk.END, content)
                    component.see(tk.END)
        except Exception as e:
            self.UI_component["progress_text"].insert(tk.END, f"Error: {e}\n")
            self.UI_component["progress_text"].see(tk.END)
        self.root.update()
        
    #Path check
    def path_check(self, source_path, destination_path):
        #check source path
        if not source_path or not Path(source_path).is_dir():
            messagebox.showerror("Error", "Please enter a valid source directory.")
            return None, None, False
        #check destination path
        if not destination_path:
            destination_path = f"{source_path}(compress)"
            Path(destination_path).mkdir(parents= True, exist_ok= True)
        if not Path(destination_path).is_dir():
            messagebox.showerror("Error", "Please enter a existed destination directory.")
            return None, None, False
        return source_path, destination_path, True

    #Check if PDF is corrupted
    def check_pdf(self,target_path):
        try:
            with open(target_path, "rb") as f:
                PdfReader(f).pages[0]
            return True, f"{Path(target_path).name} is uncorrupted"
        except Exception as e:
            return False, f"{Path(target_path).name} is corrupted: {e}"

    #Copy directories if needed
    def copy_source_dir(self, source_path, destination_path):
        for item in Path(source_path).rglob("*"):
            if item.is_dir():
                (Path(destination_path) / item.relative_to(Path(source_path))).mkdir(parents= True, exist_ok= True)

    #Compressor logic
    def ready2start(self, src_path, dest_path):
        #List all files in the source path
        total_files_list = [file for file in Path(src_path).rglob("*")
                            if file.is_file() and not file.name.startswith(("~", ".", " (1)."))]
        total_files_count = len(total_files_list)
        self.UI_component["progress_bar"]["maximum"] = total_files_count
        
        def process_file(idx, file):
            new_file_path = Path(dest_path) / Path(file).relative_to(Path(src_path))
            progress_state = f"Processed/Total files:   {idx:,}/{ total_files_count:,}"
            self.update_UI_component("progress_bar_label", f"{idx/total_files_count*100:.1f} %")
            self.update_UI_component("progress_bar", is_process_bar= True)
            
            if new_file_path.exists():
                self.update_UI_component("progress_text", f"{Path(new_file_path).name} exists. Skipping...\n")
                if new_file_path.suffix.lower() != ".pdf":
                    self.update_UI_component("not_PDF_text", f"{new_file_path.name}\n")
                self.update_UI_component("progress_state_label", progress_state)
                return 

            if Path(file).suffix.lower() == ".pdf":
                PDF_check_OK, result_PDF_check = self.check_pdf(file)
                if PDF_check_OK:
                    self.update_UI_component("progress_text", f"{Path(file).name} is processing.\n")
                    if not self.compress_pdf(file, new_file_path):
                        shutil.copy2(file, new_file_path)
                        self.update_UI_component("progress_text", f"Copied {Path(file).name} to {new_file_path} without compression.\n")
                        self.update_UI_component("progress_state_label", progress_state)
                    else: 
                        self.update_UI_component("progress_text", f"Compressed {Path(file).name} and saved to {new_file_path}.\n ")
                        PDF_check_OK, result_PDF_check = self.check_pdf(new_file_path)
                        if not PDF_check_OK:
                            self.update_UI_component("progress_text", f"{Path(file).name}: {result_PDF_check}\n")
                            self.update_UI_component("broken_text", f"{Path(file).name}")
                        else:
                            self.update_UI_component("progress_text", f"Compressed {Path(file).name} check ok.\n")
                else:
                    self.update_UI_component( "progress_text", f"{Path(file).name} is broken, didn't do anything.\n Error: {result_PDF_check}\n")
                    self.update_UI_component("broken_text", f"{Path(file).name}\n")
            else:
                shutil.copy2(file, new_file_path)
                self.update_UI_component("progress_text", f"Copied {Path(file).name} to {new_file_path} without compression.\n")
                self.update_UI_component("not_PDF_text", f"{Path(file).name}\n")
            
            self.update_UI_component("progress_state_label", progress_state)
            
        for idx, file in enumerate(total_files_list,1):
            if self.stop_compression:
                self.update_UI_component("progress_text", "Compression stopped.\n")
                break
            process_file(idx, file)
            
    #壓縮主要架構
    def start_compression(self):
        print(self.timeout)
        def onconvert():
            print(self.timeout)
            self.update_UI_component("progress_state_label", "Processed/Total file count:   0/0")
            self.update_UI_component("finish_check_label", "Completed file count:   Waiting")
            self.update_UI_component("progress_bar_label", "0 %")
            self.stop_compression= False
            
            #路徑預處理
            input_folder_path = self.UI_component["input_folder_entry"].get().strip()
            output_folder_path = self.UI_component["output_folder_entry"].get().strip()
            input_folder_path, output_folder_path, path_check_ok = self.path_check(input_folder_path, output_folder_path)                 
            
            if path_check_ok:
                self.UI_component["quit_button"].config(text="Cancel!", command= self.stop_compression_process)
                #avoid input shortly.
                self.UI_component["convert_button"].config(text= "Processing", state= tk.DISABLED)
                
                self.copy_source_dir(input_folder_path, output_folder_path)
                self.ready2start(input_folder_path, output_folder_path)

                self.update_UI_component("progress_text", "Completed")
                finish_files_list = [file for file in Path(output_folder_path).rglob("*") 
                                    if file.is_file() and not file.name.startswith(("~", ".", " (1)."))]
                self.update_UI_component("finish_check_label", f"Completed file count:   {len(finish_files_list):,}")
            
            #compression finish
            self.UI_component["quit_button"].config(text="Quit the compressor.",command= self.root.quit)
            self.UI_component["convert_button"].config(text= "Start Compression", state= tk.NORMAL)
            
        self.UI_component["convert_button"].config(command= onconvert)

    
if __name__ == "__main__":
    root1 = tk.Tk()
    root1.title("PDF compressor")
    app = PDFCompressor(root1)
    root1.mainloop()