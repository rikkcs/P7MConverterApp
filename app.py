import os
import subprocess
import threading
import urllib.request
from tkinter import Tk, Listbox, Scrollbar, messagebox, ttk, Frame, Label, Button
from pathlib import Path
import logging
import queue
from tkinterdnd2 import DND_FILES, TkinterDnD

# --- Constants ---
OPENSSL_INSTALLER_URL = "https://slproweb.com/download/Win64OpenSSL-3_3_2.exe"
APPDATA_PATH = os.getenv('APPDATA')
APP_LOG_DIR = os.path.join(APPDATA_PATH, "P7MConverterLogs")
OPENSSL_PATH = r"C:\Program Files\OpenSSL-Win64\bin\openssl.exe"
OPENSSL_INSTALLER_PATH = os.path.join(APPDATA_PATH, "Win64OpenSSL-3_3_2.exe")


# --- Setup Logging ---
os.makedirs(APP_LOG_DIR, exist_ok=True)
log_file = os.path.join(APP_LOG_DIR, 'app.log')
logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(message)s')

class OpenSSLInstaller:
    def __init__(self, update_queue):
        self.installation_window = None
        self.update_queue = update_queue

    def start_installation(self):
        self.installation_window = Tk()
        self.installation_window.title("Installing OpenSSL")
        self.installation_window.geometry("300x150")
        
        # Apply some basic styling to the installation window
        style = Style(self.installation_window)
        style.theme_use('clam')
        style.configure("TLabel", background="#f0f0f0", foreground="#333", font=("Segoe UI", 10))
        style.configure("TProgressbar", thickness=10, troughcolor="#e0e0e0", background="#007bff")
        
        Label(self.installation_window, text="Installing OpenSSL... Please wait.", style="TLabel").pack(pady=10)
        progress = Progressbar(self.installation_window, orient="horizontal", length=250, mode="indeterminate", style="TProgressbar")
        progress.pack(pady=10)
        progress.start()

        self.installation_window.update_idletasks() # Ensure window is drawn before thread starts
        threading.Thread(target=self.install_openssl).start()
        self.installation_window.mainloop() # Keep this window active

    def install_openssl(self):
        try:
            if not os.path.exists(OPENSSL_INSTALLER_PATH):
                logging.info("Downloading OpenSSL installer...")
                urllib.request.urlretrieve(OPENSSL_INSTALLER_URL, OPENSSL_INSTALLER_PATH)
            logging.info("Running OpenSSL installer...")
            install_command = f'"{OPENSSL_INSTALLER_PATH}" /verysilent /norestart'
            result = subprocess.run(install_command, shell=True, check=True, capture_output=True)
            logging.info("OpenSSL installed successfully.")
            self.update_queue.put("OpenSSL installation complete.")
        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            logging.error(f"OpenSSL installation failed: {e.stderr.decode() if hasattr(e, 'stderr') else e}")
            self.update_queue.put("OpenSSL installation failed.")
        except Exception as e:
            logging.error(f"Error during OpenSSL installation: {str(e)}")
            self.update_queue.put(f"Error installing OpenSSL: {str(e)}")
        finally:
            if self.installation_window and self.installation_window.winfo_exists(): # Check if window still exists
                self.installation_window.quit() # Stop the mainloop of the installation window
            self.update_queue.put("Launch Main App")

class P7MConverterApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("P7M to PDF Converter - Creato da Riccardo Calia 2026")
        self.geometry("900x650") # Slightly larger window
        self.minsize(600, 500)

        self.folder_path = ""
        self.files_to_convert = []
        self.update_queue = queue.Queue()

        self._create_widgets()
        self._setup_drag_and_drop()

        self.withdraw() # Hide main window until OpenSSL check is done
        self.run_check_openssl()


    def _create_widgets(self):
        main_frame = Frame(self)
        main_frame.pack(expand=True, fill="both", padx=15, pady=15)

        # --- File Selection ---
        self.select_files_button = Button(main_frame, text="Select .p7m Files", command=self.select_files)
        self.select_files_button.pack(fill="x", pady=(0, 10))

        self.select_folder_button = Button(main_frame, text="Select Folder (Recursive)", command=self.load_folder)
        self.select_folder_button.pack(fill="x", pady=(0, 15))

        # --- File Lists ---
        lists_container_frame = Frame(main_frame)
        lists_container_frame.pack(expand=True, fill="both", pady=(0, 15))
        lists_container_frame.grid_columnconfigure(0, weight=1)
        lists_container_frame.grid_columnconfigure(1, weight=1)
        lists_container_frame.grid_rowconfigure(0, weight=1)

        # To Convert List
        to_convert_frame = Frame(lists_container_frame)
        to_convert_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        Label(to_convert_frame, text="Files to Convert").pack(pady=(0, 5))

        self.file_listbox = Listbox(to_convert_frame, height=15, width=40)
        self.file_listbox.pack(expand=True, fill="both", side="left")
        file_list_scrollbar = Scrollbar(to_convert_frame, orient="vertical", command=self.file_listbox.yview)
        file_list_scrollbar.pack(side="right", fill="y")
        self.file_listbox.config(yscrollcommand=file_list_scrollbar.set)

        # Converted List
        converted_frame = Frame(lists_container_frame)
        converted_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        Label(converted_frame, text="Converted Files").pack(pady=(0, 5))

        self.converted_files_listbox = Listbox(converted_frame, height=15, width=40)
        self.converted_files_listbox.pack(expand=True, fill="both", side="left")
        converted_list_scrollbar = Scrollbar(converted_frame, orient="vertical", command=self.converted_files_listbox.yview)
        converted_list_scrollbar.pack(side="right", fill="y")
        self.converted_files_listbox.config(yscrollcommand=converted_list_scrollbar.set)

        # --- Progress and Status ---
        progress_frame = Frame(main_frame)
        progress_frame.pack(fill="x", pady=(10, 0))
        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", expand=True)
        self.status_label = Label(main_frame, text="Initializing...")
        self.status_label.pack(fill="x", pady=(5, 0))

        # --- Buttons ---
        buttons_frame = Frame(main_frame)
        buttons_frame.pack(fill="x", pady=(15, 0))

        # Use grid for buttons to center them
        buttons_frame.grid_columnconfigure(0, weight=1)
        buttons_frame.grid_columnconfigure(1, weight=1)
        buttons_frame.grid_columnconfigure(2, weight=1)
        buttons_frame.grid_columnconfigure(3, weight=1)


        self.convert_button = Button(buttons_frame, text="Convert All", command=self.start_conversion, state="disabled")
        self.convert_button.grid(row=0, column=0, padx=(0, 5), sticky="e")

        self.clear_button = Button(buttons_frame, text="Clear Lists", command=self.clear_lists)
        self.clear_button.grid(row=0, column=1, padx=(5, 0), sticky="w")

        quit_button = Button(buttons_frame, text="Quit", command=self.quit_application)
        quit_button.grid(row=0, column=2, padx=(10,0), sticky="e") # Aligned to the right

        # Separator for aesthetic
        separator = Frame(main_frame, height=2, bd=1, relief="sunken")
        separator.pack(fill="x", pady=10)

        # Credits
        credits_label = Label(main_frame, text="Creato da Riccardo Calia 2026", font=("Arial", 8))
        credits_label.pack(pady=(5, 0))




    def run_check_openssl(self):
        self.after(100, self.process_queue)
        threading.Thread(target=self.check_openssl, daemon=True).start()

    def check_openssl(self):
        if not os.path.exists(OPENSSL_PATH):
            installer = OpenSSLInstaller(self.update_queue)
            self.update_queue.put("OpenSSL not found. Starting installation...")
            installer.start_installation()
        else:
            logging.info("OpenSSL is already installed.")
            self.update_queue.put("OpenSSL found. Ready to convert.")
            self.deiconify() # Show the main window if OpenSSL is found

    def select_files(self):
        from tkinter import filedialog
        file_paths = filedialog.askopenfilenames(
            title="Select .p7m files",
            filetypes=[("P7M files", "*.p7m"), ("All files", "*.*")]
        )
        if file_paths:
            self.files_to_convert.extend(file_paths)
            self.update_file_listbox()

    def load_folder(self):
        from tkinter import filedialog
        folder_path = filedialog.askdirectory(title="Select folder containing .p7m files")
        if folder_path:
            # Clear existing folder path and files to avoid duplication if multiple folders are selected
            current_files = [f for f in self.files_to_convert if os.path.dirname(f) == self.folder_path]
            if current_files and self.folder_path != folder_path:
                # If a new folder is selected, clear files from the previous folder
                self.files_to_convert = [f for f in self.files_to_convert if os.path.dirname(f) != self.folder_path]

            self.folder_path = folder_path # Update the current folder path

            # Recursively find all .p7m files in the folder and subfolders
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    if file.lower().endswith('.p7m'):
                        full_path = os.path.join(root, file)
                        self.files_to_convert.append(full_path)

            self.update_file_listbox()

    def update_file_listbox(self):
        self.file_listbox.delete(0, 'end')
        unique_files = sorted(list(set(self.files_to_convert))) # Ensure uniqueness and sort
        self.files_to_convert = unique_files # Update internal list with unique, sorted paths

        for f_path in self.files_to_convert:
            self.file_listbox.insert('end', os.path.basename(f_path))
        
        if self.files_to_convert:
            self.status_label.config(text=f"Ready to convert {len(self.files_to_convert)} files.")
            self.convert_button.config(state="normal")
        else:
            self.status_label.config(text="No .p7m files loaded. Drag and drop to begin.")
            self.convert_button.config(state="disabled")

    def clear_lists(self):
        self.files_to_convert.clear()
        self.file_listbox.delete(0, 'end')
        self.converted_files_listbox.delete(0, 'end')
        self.progress['value'] = 0
        self.status_label.config(text="Lists cleared. Drag and drop files to start.")
        self.convert_button.config(state="disabled")

    def start_conversion(self):
        if not self.files_to_convert:
            return
        
        self.convert_button.config(state="disabled")
        self.clear_button.config(state="disabled")
        self.status_label.config(text="Converting files...")
        self.converted_files_listbox.delete(0, 'end') # Clear previous converted list
        threading.Thread(target=self.convert_files).start()

    def convert_files(self):
        total_files = len(self.files_to_convert)
        self.progress['maximum'] = total_files

        for i, p7m_file_path in enumerate(self.files_to_convert):
            # Determine the directory containing the .p7m file
            p7m_file_dir = os.path.dirname(p7m_file_path)
            # Create the output directory named "p7m-to-pdf" at the same level as the directory containing the .p7m file
            output_dir = os.path.join(p7m_file_dir, "p7m-to-pdf")

            # Create the output directory if it doesn't exist
            os.makedirs(output_dir, exist_ok=True)

            base_filename = os.path.basename(p7m_file_path)

            # Ensure the file actually ends with .p7m before attempting to replace
            if base_filename.lower().endswith('.p7m'):
                output_filename = base_filename[:-4] + '.pdf' # Remove .p7m and add .pdf
            else:
                # Fallback or error handling if somehow a non-.p7m file got in
                logging.warning(f"Skipping non-.p7m file: {base_filename}")
                self.status_label.config(text=f"Skipped non-.p7m file: {base_filename}")
                self.progress['value'] = i + 1
                self.update_idletasks()
                continue

            # Handle case where the output filename might already have .pdf extension
            if output_filename.lower().endswith('.pdf.pdf'):
                output_filename = output_filename[:-8] + '.pdf'  # Remove the extra .pdf

            output_file_path = os.path.join(output_dir, output_filename)

            try:
                self.extract_p7m_content(p7m_file_path, output_file_path)
                # Use after to update GUI from main thread
                self.after(0, lambda f=output_filename: self.converted_files_listbox.insert('end', f))
                self.after(0, lambda s=f"Converted {i + 1}/{total_files} files.": self.status_label.config(text=s))
            except Exception as e:
                self.after(0, lambda s=f"Error converting {base_filename}.": self.status_label.config(text=s))
                self.after(0, lambda msg=f"Failed to convert {base_filename}:\n{e}": messagebox.showerror("Conversion Error", msg))

            self.after(0, lambda val=i+1: self.progress.config(value=val))
            self.after(0, self.update_idletasks)


        self.after(0, lambda: self.status_label.config(text=f"Conversion complete. {total_files} files processed."))
        self.after(0, lambda: self.convert_button.config(state="normal"))
        self.after(0, lambda: self.clear_button.config(state="normal"))
        self.files_to_convert.clear() # Clear list after conversion
        self.after(0, lambda: self.file_listbox.delete(0, 'end'))

    def _setup_drag_and_drop(self):
        # Enable drag and drop for the file listbox
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        # Get the dropped file paths
        files = self.tk.splitlist(event.data)
        for file_path in files:
            if file_path.lower().endswith('.p7m'):
                if file_path not in self.files_to_convert:
                    self.files_to_convert.append(file_path)
        self.update_file_listbox()

    def extract_p7m_content(self, p7m_file_path, output_file):
        command = [
            OPENSSL_PATH, 'smime', '-verify', '-noverify', '-binary',
            '-inform', 'DER', '-in', p7m_file_path, '-out', output_file
        ]
        try:
            result = subprocess.run(command, check=True, capture_output=True, text=True, timeout=60) # Added timeout
            logging.info(f"Successfully converted {p7m_file_path}")
        except subprocess.CalledProcessError as e:
            logging.error(f"Error converting {p7m_file_path}: {e.stderr}")
            raise Exception(f"OpenSSL Error: {e.stderr.strip()}")
        except subprocess.TimeoutExpired:
            logging.error(f"OpenSSL command timed out for {p7m_file_path}")
            raise Exception("OpenSSL command timed out.")
        except Exception as e:
            logging.error(f"Unexpected error during OpenSSL call for {p7m_file_path}: {str(e)}")
            raise Exception(f"Unexpected error: {str(e)}")

    def process_queue(self):
        try:
            while True:
                message = self.update_queue.get_nowait()
                if message == "Launch Main App":
                    self.deiconify() # Show the main window
                    self.status_label.config(text="OpenSSL found. Drag and drop files to begin conversion.")
                else:
                    self.status_label.config(text=message)
                logging.info(message)
        except queue.Empty:
            pass
        finally:
            self.after(100, self.process_queue)

    def quit_application(self):
        logging.info("Application closed.")
        self.destroy()

    def run(self):
        self.mainloop()

if __name__ == "__main__":
    app = P7MConverterApp()
    app.run()