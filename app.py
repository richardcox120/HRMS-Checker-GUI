import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import re
import time
import logging
from typing import List
from molmass import Formula
import pandas as pd
import fitz  # PyMuPDF
import threading
from io import StringIO
import sys

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from main import (
    process_replacements, replace_comma_with_decimal, adjust_space_around_decimal,
    fix_floats, remove_page_numbers, remove_spaces_within_brackets, 
    remove_spaces_in_formula, transform_expressions_in_text, isotope_correct,
    protect_floats, search_hrms_with_floats, search_calcd_with_floats,
    hrms_cleanup, calc_dev_calcd_and_recalcd, 
    remove_sublists_with_missing_element1_positions_swapped,
    generate_error_dictionary, error_dictionary
)

def check_conditions(cleaned_results):
    for row in cleaned_results:
        # Check if the 8th column (index 7) is empty or contains "-0.0001" or "+0.0001"
        if row[7] not in ("", "-0.0001", "+0.0001"):
            return False
        # Check if the 7th column (index 6) as a float is less than 10
        try:
            if float(row[6]) >= 10:
                return False
        except ValueError:
            # If conversion to float fails, return False
            return False
    return True

class HRMSChecker:
    def __init__(self):
        self.setup_gui()
        
    def setup_gui(self):
        # Create main window
        self.root = tk.Tk()
        self.root.title('HRMS Checker 2.0')
        self.root.geometry('800x700')
        
        # Title
        title_frame = tk.Frame(self.root)
        title_frame.pack(pady=10)
        tk.Label(title_frame, text='HRMS Checker 2.0', font=('Arial', 16, 'bold')).pack()
        
        # Separator
        ttk.Separator(self.root, orient='horizontal').pack(fill='x', pady=5)
        
        # Input section
        input_frame = tk.Frame(self.root)
        input_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(input_frame, text='Select source folder containing PDF files:').pack(anchor='w')
        
        folder_frame = tk.Frame(input_frame)
        folder_frame.pack(fill='x', pady=5)
        
        self.folder_var = tk.StringVar()
        self.folder_entry = tk.Entry(folder_frame, textvariable=self.folder_var, width=60)
        self.folder_entry.pack(side='left', fill='x', expand=True)
        
        self.browse_button = tk.Button(folder_frame, text='Browse', command=self.browse_folder)
        self.browse_button.pack(side='right', padx=(5, 0))
        
        # Report checkbox
        self.report_var = tk.BooleanVar(value=True)
        tk.Checkbutton(input_frame, text='Generate Excel report', variable=self.report_var).pack(anchor='w', pady=5)
        
        # Control buttons
        button_frame = tk.Frame(self.root)
        button_frame.pack(fill='x', padx=10, pady=5)
        
        self.start_button = tk.Button(button_frame, text='Start Analysis', command=self.start_analysis, 
                                     bg='green', fg='white', font=('Arial', 10, 'bold'))
        self.start_button.pack(side='left', padx=(0, 5))
        
        self.clear_button = tk.Button(button_frame, text='Clear Results', command=self.clear_results)
        self.clear_button.pack(side='left', padx=(0, 5))
        
        tk.Button(button_frame, text='Exit', command=self.root.quit, bg='red', fg='white').pack(side='right')
        
        # Separator
        ttk.Separator(self.root, orient='horizontal').pack(fill='x', pady=5)
        
        # Progress section
        progress_frame = tk.Frame(self.root)
        progress_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(progress_frame, text='Progress:').pack(anchor='w')
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.pack(fill='x', pady=2)
        
        self.status_label = tk.Label(progress_frame, text='Status: Ready', anchor='w')
        self.status_label.pack(fill='x')
        
        # Separator
        ttk.Separator(self.root, orient='horizontal').pack(fill='x', pady=5)
        
        # Results section
        results_frame = tk.Frame(self.root)
        results_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        tk.Label(results_frame, text='Results:', font=('Arial', 12, 'bold')).pack(anchor='w')
        
        self.results_text = scrolledtext.ScrolledText(results_frame, wrap=tk.NONE, 
                                                     font=('Courier', 9), height=20)
        self.results_text.pack(fill='both', expand=True)
        
        # Summary section
        summary_frame = tk.Frame(self.root)
        summary_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(summary_frame, text='Summary:', font=('Arial', 12, 'bold')).pack(anchor='w')
        
        summary_info_frame = tk.Frame(summary_frame)
        summary_info_frame.pack(fill='x')
        
        self.total_label = tk.Label(summary_info_frame, text='Total measurements: 0')
        self.total_label.pack(side='left')
        
        self.time_label = tk.Label(summary_info_frame, text='Processing time: 0s')
        self.time_label.pack(side='right')
        
    def browse_folder(self):
        """Open folder browser dialog"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_var.set(folder_path)
    
    def update_progress(self, current, total, status_text):
        """Update progress bar and status text"""
        if total > 0:
            progress = int((current / total) * 100)
            self.progress_bar['value'] = progress
        self.status_label.config(text=f'Status: {status_text}')
        self.root.update_idletasks()
    
    def log_to_results(self, text):
        """Add text to results area"""
        self.results_text.insert(tk.END, text)
        self.results_text.see(tk.END)
        self.root.update_idletasks()
    
    def show_error(self, message):
        """Show error message"""
        messagebox.showerror('Error', message)
        self.log_to_results(f"ERROR: {message}\n")
        self.status_label.config(text=f'Status: Error - {message}')
    
    def start_analysis(self):
        """Start the PDF analysis"""
        folder_path = self.folder_var.get()
        if not folder_path:
            self.show_error('Please select a source folder first!')
            return
        
        if not os.path.exists(folder_path):
            self.show_error('Selected folder does not exist!')
            return
        
        write_report = self.report_var.get()
        
        # Disable buttons during processing
        self.start_button.config(state='disabled')
        self.clear_button.config(state='disabled')
        self.browse_button.config(state='disabled')
        
        # Clear previous results
        self.results_text.delete(1.0, tk.END)
        self.log_to_results(f"Starting analysis of folder: {folder_path}\n")
        
        # Start processing in separate thread
        thread = threading.Thread(
            target=self.process_pdfs_thread,
            args=(folder_path, write_report),
            daemon=True
        )
        thread.start()
    
    def process_pdfs_thread(self, folder_path, write_report):
        """Process PDFs in a separate thread"""
        try:
            start_time = time.time()
            hrms_total_measurements = 0
            
            # Get list of PDF files
            pdf_files = self.list_pdfs_in_folder(folder_path)
            total_files = len(pdf_files)
            
            if total_files == 0:
                self.log_to_results("No PDF files found to process.\n")
                self.update_progress(0, 100, "No files found")
                return
            
            destination_folder = os.path.join(folder_path, 'HRMS_report')
            if write_report and not os.path.exists(destination_folder):
                os.makedirs(destination_folder)
            
            # Process each PDF file
            for i, pdf_file_path in enumerate(pdf_files):
                filename = os.path.basename(pdf_file_path)
                self.update_progress(i, total_files, f"Processing {i+1}/{total_files}: {filename}")
                
                # Extract and process text
                text_content = self.extract_text_from_pdf(pdf_file_path)
                if not text_content:
                    self.log_to_results(f"Warning: No text extracted from {filename}\n")
                    continue
                
                # Apply all text processing
                processed_text = self.process_text_content(text_content)
                
                # Extract HRMS data
                results = self.extract_hrms_data(processed_text)
                cleaned_results = self.cleanup_hrms_data(results)
                
                num_measurements = len(cleaned_results)
                hrms_total_measurements += num_measurements
                
                if cleaned_results:
                    self.log_to_results(f"\n{pdf_file_path}\n")
                    self.display_results_table(cleaned_results, pdf_file_path, 
                                             write_report, destination_folder)
                    
                    if check_conditions(cleaned_results):
                        self.log_to_results("Awesome! No mistakes!\n")
                else:
                    self.log_to_results(f"\nNo HRMS matches found in {filename}\n")
            
            # Final summary
            elapsed_time = time.time() - start_time
            minutes, seconds = divmod(elapsed_time, 60)
            
            self.log_to_results(f"\nTotal measurements found: {hrms_total_measurements}\n")
            self.log_to_results(f"Processing completed in {int(minutes)}m {int(seconds)}s\n")
            
            # Update summary in GUI
            self.total_label.config(text=f'Total measurements: {hrms_total_measurements}')
            self.time_label.config(text=f'Processing time: {int(minutes)}m {int(seconds)}s')
            self.update_progress(100, 100, "Analysis complete!")
            
        except Exception as e:
            self.log_to_results(f"Error during processing: {str(e)}\n")
            self.update_progress(0, 100, "Error occurred")
        
        finally:
            # Re-enable buttons
            self.start_button.config(state='normal')
            self.clear_button.config(state='normal')
            self.browse_button.config(state='normal')
    
    def clear_results(self):
        """Clear all results and reset the interface"""
        self.results_text.delete(1.0, tk.END)
        self.progress_bar['value'] = 0
        self.status_label.config(text='Status: Ready')
        self.total_label.config(text='Total measurements: 0')
        self.time_label.config(text='Processing time: 0s')
    
    def list_pdfs_in_folder(self, directory_path):
        """List all PDF files in directory"""
        try:
            return [os.path.join(directory_path, f) for f in os.listdir(directory_path) 
                   if f.lower().endswith('.pdf')]
        except Exception as e:
            self.log_to_results(f"Error listing PDFs: {e}\n")
            return []
    
    def extract_text_from_pdf(self, file_path):
        """Extract text from PDF using PyMuPDF"""
        try:
            with fitz.open(file_path) as pdf_document:
                text_content = ""
                for page_num in range(pdf_document.page_count):
                    page = pdf_document.load_page(page_num)
                    text_content += page.get_text()
                return text_content
        except Exception as e:
            self.log_to_results(f"Error extracting text from {file_path}: {e}\n")
            return ""
    
    def process_text_content(self, text_content):
        """Apply all text processing steps using main.py functions"""
        # Apply all the processing steps from your main.py
        text_content = re.sub(r'\s+', ' ', text_content).strip()
        text_content = process_replacements(text_content)
        text_content = replace_comma_with_decimal(text_content)
        text_content = adjust_space_around_decimal(text_content)
        text_content = fix_floats(text_content)
        text_content = remove_page_numbers(text_content)
        text_content = re.sub(r'\[((C\d+(?:[A-Z][a-z]?\d*)*),\s*([M+][^]]+))', r'\1 [\3]', text_content)
        text_content = re.sub(r'(C)(\d+)(h)(\d+)', lambda m: f'C{m.group(2)}H{m.group(4)}', text_content, flags=re.IGNORECASE)
        text_content = re.sub(r'(c)(\d+)(H)(\d+)', lambda m: f'C{m.group(2)}H{m.group(4)}', text_content, flags=re.IGNORECASE)
        text_content = re.sub(r'\b(C)(\d+)(HD)\b', r'C\2H1D', text_content)
        text_content = re.sub(r'\b(C)\s*(\d*)\s*(H)\s*(\d*)\s*(N)\s*(\d*)\b',
                      lambda m: f"{m.group(1)}{m.group(2) or ''}{m.group(3)}{m.group(4) or ''}{m.group(5)}{m.group(6) or ''}",
                      text_content)
        text_content = re.sub(r'\b(C)\s*(\d*)\s*(H)\s*(\d*)\s*(O)\s*(\d*)\b',
                      lambda m: f"{m.group(1)}{m.group(2) or ''}{m.group(3)}{m.group(4) or ''}{m.group(5)}{m.group(6) or ''}",
                      text_content)
        
        text_content = text_content.replace("C2o","C20").replace("C1o","C10").replace("Cal","cal")
        text_content = re.sub(r'B(\d+)H(\d+)', r'H\2B\1', text_content)
        text_content = text_content.replace('\n', ' ').replace('+-', '+').replace(':'," ").replace('–','-').replace(','," ")
        text_content = remove_spaces_within_brackets(text_content)
        
        # Remove nested brackets
        text_content = re.sub(r'\(\[([^]]{1,10})]\+\)', r'[\1]+', text_content)
        text_content = re.sub(r'\[\[([^]]{1,10})]\+]', r'[\1]+', text_content)
        text_content = text_content.replace(' [[', '[').replace(']]', ']')
        
        # Apply replacements
        replacements = {
            "₁": "1", "₂": "2", "₃": "3", "₄": "4", "₅": "5",
            "₆": "6", "₇": "7", "₈": "8", "₉": "9", "₀": "0", "¹": "1", "²": "2", "³": "3",
            "⁴": "4", "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9", "⁰": "0","С":"C","Н":"H",
            "C ": "C", " H ": "H", " F ":"F", " N ": "N", " Cl ":"Cl", " Br ":"Br", " O ": "O"," I ": "I",
            " P ":"P"," B ":"B", " S ":"S"," NO ":"NO", " Na ": "Na", " SNa ": "SNa"," NNa ":"NNa",
            " + ":"+ ",
        }
        
        for original, replacement in replacements.items():
            text_content = text_content.replace(original, replacement)
            
        text_content = remove_spaces_in_formula(text_content)
        text_content = text_content.replace('#', '')
        text_content = re.sub(r'(C\d+)', r' \1', text_content)
        text_content = transform_expressions_in_text(text_content)
        text_content = isotope_correct(text_content)
        text_content = protect_floats(text_content)
        text_content = text_content.replace("[13C]","H1HeXe")
        text_content = text_content.replace("CF", "C1F")
        text_content = text_content.replace("HN", "H1N")
        
        return text_content
    
    def extract_hrms_data(self, text_content):
        """Extract HRMS data from processed text using main.py functions"""
        results1 = search_hrms_with_floats(text_content)
        modified_text = text_content
        for match in results1:
            modified_text = modified_text.replace(match, '')
        
        # Clean up any double spaces created by the removals
        modified_text = re.sub(r'\s+', ' ', modified_text).strip()
        text_content = modified_text
        results2 = search_calcd_with_floats(text_content)
        
        return results1 + results2
    
    def cleanup_hrms_data(self, results):
        """Clean up and process HRMS data using main.py functions"""
        cleaned_results = hrms_cleanup(results, error_dictionary)
        cleaned_results = calc_dev_calcd_and_recalcd(cleaned_results)
        cleaned_results = remove_sublists_with_missing_element1_positions_swapped(cleaned_results)
        
        # Remove duplicates while preserving order
        cleaned_results_new = []
        for sublist in cleaned_results:
            if sublist not in cleaned_results_new:
                cleaned_results_new.append(sublist)
        
        return cleaned_results_new
    
    def display_results_table(self, cleaned_results, pdf_file_path, write_report, destination_folder):
        """Display results in a formatted table"""
        headers = ['Formula', 'Ion', 'Calcd Mass', 'Found Mass', 'Recalcd Mass', 
                  'Dev (Calcd)', 'Dev (Recalcd)', 'Potential error']
        
        # Calculate column widths
        col_widths = [max(len(str(row[i])) for row in [headers] + cleaned_results) 
                     for i in range(8)]
        
        # Print headers
        header_row = '  '.join(f"{headers[i]:<{col_widths[i]}}" for i in range(8))
        self.log_to_results(header_row + '\n')
        self.log_to_results('-' * len(header_row) + '\n')
        
        # Print data rows
        excel_data = []
        for row in cleaned_results:
            # Format row for display
            row_output = []
            excel_row = []
            
            for i in range(8):
                cell_content = row[i]
                excel_row.append(cell_content)
                formatted_cell = f"{cell_content:<{col_widths[i]}}"
                row_output.append(formatted_cell)
            
            self.log_to_results('  '.join(row_output) + '\n')
            excel_data.append(excel_row)
        
        if write_report and excel_data:
            try:
                base_filename = os.path.basename(pdf_file_path)
                filename_without_ext = os.path.splitext(base_filename)[0]
                output_filename = f"output {filename_without_ext}.xlsx"
                excel_file_path = os.path.join(destination_folder, output_filename)
                
                df = pd.DataFrame(excel_data, columns=headers)
                df.to_excel(excel_file_path, index=False)
                self.log_to_results(f"Excel report saved: {excel_file_path}\n")
            except Exception as e:
                self.log_to_results(f"Error saving Excel file: {e}\n")
    
    def run(self):
        """Start the GUI application"""
        self.root.mainloop()

def main():
    """Main function to start the GUI application"""
    try:
        app = HRMSChecker()
        app.run()
    except Exception as e:
        print(f'Application error: {str(e)}')

if __name__ == '__main__':
    main()