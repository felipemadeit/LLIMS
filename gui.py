import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from Read.excel_chain_data_reader import excel_chain_data_reader
from Read.matrix_data_reader import matrix_data_reader
from openpyxl import load_workbook
import os

class SR_LIMS_App:
    """
    SR-LIMS (Scientific Research Laboratory Information Management System)
    Excel Data Processing Application
    
    Features:
    - Excel file selection with validation
    - Chain data and matrix data processing
    - Progress feedback
    - Professional UI with themed widgets
    """
    
    def __init__(self, root):
        self.root = root
        self.file_path = tk.StringVar()
        self.progress = tk.IntVar()
        self.root.title("SR-LIMS v1.0")
        self.root.geometry("800x600")
        self.setup_ui()
        
        # Configure styles
        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TLabel', background='#f0f0f0', font=('Helvetica', 10))
        self.style.configure('TButton', font=('Helvetica', 10))
        self.style.configure('Header.TLabel', font=('Helvetica', 14, 'bold'))
        
        # Initialize variables
        self.file_path = tk.StringVar()
        self.progress = tk.IntVar()
        self.progress.set(0)
        
    def setup_ui(self):
        """Initialize all UI components"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header = ttk.Label(main_frame, text="SR-LIMS Data Processor", style='Header.TLabel')
        header.pack(pady=(0, 20))
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="Excel File Selection", padding=10)
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(file_frame, text="Select Excel File:").grid(row=0, column=0, sticky=tk.W)
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=50)
        file_entry.grid(row=0, column=1, padx=5)
        
        browse_btn = ttk.Button(file_frame, text="Browse...", command=self.browse_file)
        browse_btn.grid(row=0, column=2, padx=5)
        
        # Processing frame
        process_frame = ttk.LabelFrame(main_frame, text="Data Processing", padding=10)
        process_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_text = tk.Text(process_frame, height=15, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Progress bar
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(progress_frame, text="Progress:").pack(side=tk.LEFT)
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            orient=tk.HORIZONTAL, 
            length=400, 
            mode='determinate',
            variable=self.progress
        )
        self.progress_bar.pack(side=tk.LEFT, padx=5)
        
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        process_btn = ttk.Button(
            button_frame, 
            text="Process Data", 
            command=self.process_data,
            style='Accent.TButton'
        )
        process_btn.pack(side=tk.RIGHT, padx=5)
        
        exit_btn = ttk.Button(button_frame, text="Exit", command=self.root.quit)
        exit_btn.pack(side=tk.RIGHT)
        
        # Configure grid weights
        file_frame.columnconfigure(1, weight=1)
        
    def browse_file(self):
        """Open file dialog to select Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        
        if file_path:
            self.file_path.set(file_path)
            self.log_message(f"Selected file: {file_path}")
    
    def process_data(self):
        """Process the selected Excel file"""
        file_path = self.file_path.get()
        
        if not file_path:
            messagebox.showwarning("Warning", "Please select an Excel file first!")
            return
            
        if not os.path.exists(file_path):
            messagebox.showerror("Error", "The selected file does not exist!")
            return
            
        try:
            self.log_message("\n=== Starting Data Processing ===")
            self.log_message("Loading workbook...")
            self.update_progress(10)
            
            # Load workbook with formulas calculated
            wb_to_read = load_workbook(filename=file_path, data_only=True)
            self.log_message("Workbook loaded successfully")
            self.update_progress(20)
            
            # Read chain data
            self.log_message("\nReading chain of custody data...")
            chain_data = excel_chain_data_reader(wb_to_read, file_path, [23, 23])
            self.log_message(f"Found {len(chain_data)} samples in chain data")
            self.update_progress(50)
            
            # Read matrix data
            self.log_message("\nProcessing matrix data...")
            matrix_data_reader(wb_to_read, chain_data)
            self.log_message("Matrix data processing completed")
            self.update_progress(90)
            
            # Completion
            self.log_message("\n=== Processing Complete ===")
            self.update_progress(100)
            messagebox.showinfo("Success", "Data processing completed successfully!")
            
        except Exception as e:
            self.log_message(f"\nERROR: {str(e)}", error=True)
            messagebox.showerror("Processing Error", f"An error occurred:\n{str(e)}")
            self.progress.set(0)
    
    def log_message(self, message, error=False):
        """Add message to log text area"""
        self.log_text.config(state=tk.NORMAL)
        
        if error:
            self.log_text.insert(tk.END, message + "\n", "error")
        else:
            self.log_text.insert(tk.END, message + "\n")
            
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)
        
        # Force UI update
        self.root.update_idletasks()
    
    def update_progress(self, value):
        """Update progress bar value"""
        self.progress.set(value)
        self.root.update_idletasks()

def main():
    root = tk.Tk()
    
    # Set theme (requires ttkthemes or use system theme)
    try:
        from ttkthemes import ThemedStyle
        style = ThemedStyle(root)
        style.set_theme("arc")
    except ImportError:
        pass
    
    # Configure text tags for coloring
    root.option_add('*TCombobox*Listbox.font', ('Helvetica', 10))
    root.option_add('*TCombobox*Listbox.background', 'white')
    
    app = SR_LIMS_App(root)
    
    # Create text tags for styling
    app.log_text.tag_config("error", foreground="red")
    
    root.mainloop()

if __name__ == "__main__":
    main()