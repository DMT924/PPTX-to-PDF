import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
import win32com.client
import threading
import os
import webbrowser

class ModernUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title("PPTX to PDF Converter")
        self.minsize(600, 450)
        
        # Center window on screen
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = 600
        window_height = 450
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f'{window_width}x{window_height}+{x}+{y}')
        
        # Set theme
        ctk.set_appearance_mode("dark")  # Can be "dark" or "light"
        ctk.set_default_color_theme("blue")

        # Create main container
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(expand=True, fill='both', padx=20, pady=20)

        # Title
        title_label = ctk.CTkLabel(
            self.main_frame,
            text="PowerPoint to PDF Converter",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(0, 20))

        # Mode selection
        self.folder_var = tk.BooleanVar(value=False)
        mode_frame = ctk.CTkFrame(self.main_frame)
        mode_frame.pack(fill='x', pady=(0, 15))

        mode_label = ctk.CTkLabel(
            mode_frame,
            text="Convert mode:",
            font=ctk.CTkFont(size=14)
        )
        mode_label.pack(side='left', padx=(0, 10))

        single_radio = ctk.CTkRadioButton(
            mode_frame,
            text="Single File",
            variable=self.folder_var,
            value=False
        )
        single_radio.pack(side='left', padx=10)

        folder_radio = ctk.CTkRadioButton(
            mode_frame,
            text="Folder",
            variable=self.folder_var,
            value=True
        )
        folder_radio.pack(side='left', padx=10)

        # File selection
        select_frame = ctk.CTkFrame(self.main_frame)
        select_frame.pack(fill='x', pady=(0, 15))

        self.file_entry = ctk.CTkEntry(
            select_frame,
            placeholder_text="Select a PowerPoint file or folder..."
        )
        self.file_entry.pack(side='left', expand=True, fill='x', padx=(0, 10))

        browse_button = ctk.CTkButton(
            select_frame,
            text="Browse",
            command=self.select_path,
            width=100
        )
        browse_button.pack(side='right')

        # Status label
        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="",
            font=ctk.CTkFont(size=12)
        )
        self.status_label.pack(pady=(10, 5))

        # Progress bar
        self.progress = ctk.CTkProgressBar(
            self.main_frame,
            width=400,
            height=10
        )
        self.progress.pack(pady=15)
        self.progress.set(0)

        # Convert button
        self.convert_button = ctk.CTkButton(
            self.main_frame,
            text="Convert to PDF",
            command=self.convert_file,
            font=ctk.CTkFont(size=15, weight="bold"),
            height=40
        )
        self.convert_button.pack(pady=20)

    def select_path(self):
        """Handle file/folder selection based on mode"""
        if self.folder_var.get():
            # Folder mode
            path = filedialog.askdirectory()
        else:
            # Single file mode
            path = filedialog.askopenfilename(
                filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")]
            )
            
        if path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, path)

    def convert_to_pdf(self, pptx_path, pdf_path):
        """Convert PowerPoint to PDF"""
        powerpoint = None
        deck = None
        try:
            powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
            powerpoint.Visible = 0
            powerpoint.DisplayAlerts = 0
            
            # Convert paths to absolute paths
            pptx_path = os.path.abspath(pptx_path)
            pdf_path = os.path.abspath(pdf_path)
            
            self.status_label.config(text=f"Opening: {pptx_path}")
            deck = powerpoint.Presentations.Open(pptx_path, ReadOnly=True, WithWindow=False)
            
            self.status_label.config(text=f"Saving as: {pdf_path}")
            deck.SaveAs(pdf_path, 32)  # 32 is the PDF format code
            
            return True
            
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")
            return False
            
        finally:
            if deck:
                try:
                    deck.Close()
                except:
                    pass
            if powerpoint:
                try:
                    powerpoint.Quit()
                except:
                    pass

    def process_conversion(self, path):
        try:
            self.status_label.config(text="Starting conversion...")
            self.convert_button.config(state=tk.DISABLED)
            
            if os.path.isdir(path):
                # Process directory
                pptx_files = [f for f in os.listdir(path) if f.endswith('.pptx')]
                total_files = len(pptx_files)
                
                if total_files == 0:
                    messagebox.showwarning("No Files Found", "No PowerPoint files found in the selected folder.")
                    return
                
                # Create output folder
                output_folder = os.path.join(path, "PDF_Conversions")
                os.makedirs(output_folder, exist_ok=True)
                
                converted_count = 0
                failed_files = []
                
                for i, pptx_file in enumerate(pptx_files):
                    pptx_path = os.path.join(path, pptx_file)
                    pdf_name = os.path.splitext(pptx_file)[0] + '.pdf'
                    pdf_path = os.path.join(output_folder, pdf_name)
                    
                    self.status_label.config(text=f"Converting {i+1}/{total_files}: {pptx_file}")
                    self.progress['value'] = (i / total_files) * 100
                    self.main_frame.update_idletasks()
                    
                    if self.convert_to_pdf(pptx_path, pdf_path):
                        converted_count += 1
                    else:
                        failed_files.append(pptx_file)
                
                # Show summary
                summary = f"Converted {converted_count} of {total_files} files.\nOutput folder: {output_folder}"
                if failed_files:
                    summary += f"\n\nFailed files:\n" + "\n".join(failed_files)
                
                messagebox.showinfo("Conversion Complete", summary)
                
                # Add button to open output folder
                folder_button = tk.Button(self.main_frame, text="Open Output Folder", 
                                        command=lambda: webbrowser.open(output_folder))
                folder_button.pack(pady=5)
                
            else:
                # Process single file
                pdf_path = os.path.splitext(path)[0] + '.pdf'
                
                if os.path.exists(pdf_path):
                    if not messagebox.askyesno("File exists", "PDF already exists. Do you want to overwrite?"):
                        return
                
                self.progress['value'] = 50
                success = self.convert_to_pdf(path, pdf_path)
                self.progress['value'] = 100
                
                if success and os.path.exists(pdf_path):
                    preview_button = tk.Button(self.main_frame, text="Open PDF", 
                                             command=lambda: webbrowser.open(pdf_path))
                    preview_button.pack()
                    messagebox.showinfo("Success", f"PDF saved at {pdf_path}")
                else:
                    messagebox.showerror("Error", "Conversion failed")
                
        except Exception as e:
            messagebox.showerror("Error", f"Conversion error: {str(e)}")
        finally:
            self.convert_button.config(state=tk.NORMAL)
            self.progress['value'] = 0
            self.status_label.config(text="")

    def convert_file(self):
        path = self.file_entry.get()
        if not path:
            messagebox.showerror("Error", "Please select a file or folder.")
            return
            
        if os.path.isdir(path) or path.endswith('.pptx'):
            threading.Thread(target=self.process_conversion, args=(path,)).start()
        else:
            messagebox.showerror("Invalid Selection", "Please select a .pptx file or a folder containing .pptx files.")

# Create the main window
root = ModernUI()
root.mainloop()