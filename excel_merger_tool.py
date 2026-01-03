import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
from datetime import datetime
import threading

class ExcelProcessorTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Merger & Processor - by Dumok Data Lab")
        self.root.geometry("1000x750")
        self.root.configure(bg='#1a1a2e')
        
        # 스타일
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'), background='#1a1a2e', foreground='white')
        style.configure('TLabel', background='#1a1a2e', foreground='white', font=('Arial', 10))
        
        self.files = []
        self.merged_df = None
        
        self.create_widgets()
        
    def create_widgets(self):
        # 헤더
        header_frame = tk.Frame(self.root, bg='#16213e', pady=20)
        header_frame.pack(fill='x')
        
        title = ttk.Label(header_frame, text="📊 Excel Merger & Processor", style='Title.TLabel')
        title.pack()
        
        subtitle = ttk.Label(header_frame, text="Merge, filter, and process multiple Excel files instantly", 
                           font=('Arial', 10))
        subtitle.pack()
        
        # 메인 컨테이너
        main_frame = tk.Frame(self.root, bg='#1a1a2e', padx=25, pady=20)
        main_frame.pack(fill='both', expand=True)
        
        # 좌측: 파일 관리
        left_frame = tk.Frame(main_frame, bg='#1a1a2e')
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 15))
        
        # 파일 선택 섹션
        file_section = tk.LabelFrame(left_frame, text=" 📁 Select Files ", 
                                    bg='#16213e', fg='white', font=('Arial', 11, 'bold'),
                                    padx=15, pady=15)
        file_section.pack(fill='both', expand=True, pady=(0, 15))
        
        btn_frame = tk.Frame(file_section, bg='#16213e')
        btn_frame.pack(fill='x', pady=(0, 10))
        
        self.add_btn = tk.Button(btn_frame, text="➕ Add Files", command=self.add_files,
                                bg='#0f3460', fg='white', font=('Arial', 10, 'bold'),
                                relief='flat', padx=15, pady=8, cursor='hand2')
        self.add_btn.pack(side='left', padx=(0, 10))
        
        self.clear_btn = tk.Button(btn_frame, text="🗑️ Clear All", command=self.clear_files,
                                  bg='#e94560', fg='white', font=('Arial', 10),
                                  relief='flat', padx=15, pady=8, cursor='hand2')
        self.clear_btn.pack(side='left')
        
        # 파일 리스트
        list_frame = tk.Frame(file_section, bg='#16213e')
        list_frame.pack(fill='both', expand=True)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')
        
        self.file_listbox = tk.Listbox(list_frame, font=('Consolas', 9),
                                       bg='#0f3460', fg='white',
                                       selectbackground='#e94560',
                                       yscrollcommand=scrollbar.set,
                                       relief='flat', borderwidth=0)
        self.file_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        self.file_count_label = ttk.Label(file_section, text="Files: 0", font=('Arial', 9))
        self.file_count_label.pack(anchor='w', pady=(10, 0))
        
        # 우측: 처리 옵션 및 실행
        right_frame = tk.Frame(main_frame, bg='#1a1a2e', width=350)
        right_frame.pack(side='right', fill='both')
        right_frame.pack_propagate(False)
        
        # 처리 옵션
        options_section = tk.LabelFrame(right_frame, text=" ⚙️ Processing Options ",
                                       bg='#16213e', fg='white', font=('Arial', 11, 'bold'),
                                       padx=15, pady=15)
        options_section.pack(fill='x', pady=(0, 15))
        
        self.merge_var = tk.BooleanVar(value=True)
        self.remove_duplicates_var = tk.BooleanVar(value=False)
        self.add_source_var = tk.BooleanVar(value=True)
        
        tk.Checkbutton(options_section, text="Merge all sheets", variable=self.merge_var,
                      bg='#16213e', fg='white', selectcolor='#0f3460',
                      font=('Arial', 10), activebackground='#16213e',
                      activeforeground='white').pack(anchor='w', pady=5)
        
        tk.Checkbutton(options_section, text="Remove duplicates", variable=self.remove_duplicates_var,
                      bg='#16213e', fg='white', selectcolor='#0f3460',
                      font=('Arial', 10), activebackground='#16213e',
                      activeforeground='white').pack(anchor='w', pady=5)
        
        tk.Checkbutton(options_section, text="Add source filename column", variable=self.add_source_var,
                      bg='#16213e', fg='white', selectcolor='#0f3460',
                      font=('Arial', 10), activebackground='#16213e',
                      activeforeground='white').pack(anchor='w', pady=5)
        
        # 필터 섹션
        filter_section = tk.LabelFrame(right_frame, text=" 🔍 Filter (Optional) ",
                                      bg='#16213e', fg='white', font=('Arial', 11, 'bold'),
                                      padx=15, pady=15)
        filter_section.pack(fill='x', pady=(0, 15))
        
        ttk.Label(filter_section, text="Column name:").pack(anchor='w')
        self.filter_col_entry = tk.Entry(filter_section, font=('Arial', 10),
                                        bg='#0f3460', fg='white', relief='flat',
                                        insertbackground='white')
        self.filter_col_entry.pack(fill='x', pady=(5, 10), ipady=6)
        
        ttk.Label(filter_section, text="Contains text:").pack(anchor='w')
        self.filter_text_entry = tk.Entry(filter_section, font=('Arial', 10),
                                         bg='#0f3460', fg='white', relief='flat',
                                         insertbackground='white')
        self.filter_text_entry.pack(fill='x', pady=(5, 0), ipady=6)
        
        # 실행 버튼
        action_frame = tk.Frame(right_frame, bg='#1a1a2e')
        action_frame.pack(fill='x', pady=(0, 15))
        
        self.process_btn = tk.Button(action_frame, text="▶️ Process Files",
                                     command=self.process_files,
                                     bg='#00d9ff', fg='#1a1a2e', font=('Arial', 12, 'bold'),
                                     relief='flat', pady=12, cursor='hand2',
                                     state='disabled')
        self.process_btn.pack(fill='x')
        
        self.export_btn = tk.Button(action_frame, text="💾 Export Result",
                                    command=self.export_result,
                                    bg='#4caf50', fg='white', font=('Arial', 11),
                                    relief='flat', pady=10, cursor='hand2',
                                    state='disabled')
        self.export_btn.pack(fill='x', pady=(10, 0))
        
        # 통계 표시
        stats_section = tk.LabelFrame(right_frame, text=" 📈 Statistics ",
                                     bg='#16213e', fg='white', font=('Arial', 11, 'bold'),
                                     padx=15, pady=15)
        stats_section.pack(fill='both', expand=True)
        
        self.stats_text = tk.Text(stats_section, font=('Consolas', 9),
                                 bg='#0f3460', fg='#00d9ff',
                                 relief='flat', height=8, wrap='word')
        self.stats_text.pack(fill='both', expand=True)
        self.stats_text.insert(1.0, "No data processed yet")
        self.stats_text.config(state='disabled')
        
        # 하단: 로그
        log_frame = tk.LabelFrame(self.root, text=" 📋 Processing Log ",
                                 bg='#16213e', fg='white', font=('Arial', 10, 'bold'),
                                 padx=10, pady=10)
        log_frame.pack(fill='both', expand=True, padx=25, pady=(0, 20))
        
        self.log_text = scrolledtext.ScrolledText(log_frame, font=('Consolas', 9),
                                                  bg='#0f3460', fg='#ffffff',
                                                  relief='flat', height=8)
        self.log_text.pack(fill='both', expand=True)
        self.log("Welcome to Excel Processor Tool!")
        self.log("Add Excel files to begin...")
        
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        for file in files:
            if file not in self.files:
                self.files.append(file)
                self.file_listbox.insert(tk.END, os.path.basename(file))
        
        self.file_count_label.config(text=f"Files: {len(self.files)}")
        self.process_btn.config(state='normal' if self.files else 'disabled')
        self.log(f"Added {len(files)} file(s)")
        
    def clear_files(self):
        self.files.clear()
        self.file_listbox.delete(0, tk.END)
        self.file_count_label.config(text="Files: 0")
        self.process_btn.config(state='disabled')
        self.export_btn.config(state='disabled')
        self.merged_df = None
        self.log("Cleared all files")
        
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        
    def process_files(self):
        if not self.files:
            messagebox.showwarning("No Files", "Please add Excel files first")
            return
        
        self.process_btn.config(state='disabled', text="⏳ Processing...")
        self.log("=" * 50)
        self.log("Starting file processing...")
        
        thread = threading.Thread(target=self.process_thread)
        thread.daemon = True
        thread.start()
        
    def process_thread(self):
        try:
            dfs = []
            
            for idx, file in enumerate(self.files, 1):
                self.log(f"Reading file {idx}/{len(self.files)}: {os.path.basename(file)}")
                
                try:
                    df = pd.read_excel(file)
                    
                    if self.add_source_var.get():
                        df['Source_File'] = os.path.basename(file)
                    
                    dfs.append(df)
                    self.log(f"  ✓ Loaded {len(df)} rows")
                    
                except Exception as e:
                    self.log(f"  ✗ Error reading file: {str(e)}")
            
            if not dfs:
                self.log("No data loaded")
                return
            
            # 병합
            self.log("Merging data...")
            self.merged_df = pd.concat(dfs, ignore_index=True)
            self.log(f"  ✓ Merged into {len(self.merged_df)} total rows")
            
            # 중복 제거
            if self.remove_duplicates_var.get():
                before = len(self.merged_df)
                self.merged_df = self.merged_df.drop_duplicates()
                removed = before - len(self.merged_df)
                self.log(f"  ✓ Removed {removed} duplicate rows")
            
            # 필터 적용
            filter_col = self.filter_col_entry.get().strip()
            filter_text = self.filter_text_entry.get().strip()
            
            if filter_col and filter_text:
                if filter_col in self.merged_df.columns:
                    before = len(self.merged_df)
                    self.merged_df = self.merged_df[
                        self.merged_df[filter_col].astype(str).str.contains(filter_text, case=False, na=False)
                    ]
                    filtered = before - len(self.merged_df)
                    self.log(f"  ✓ Applied filter: {before - filtered} rows remaining")
                else:
                    self.log(f"  ⚠ Column '{filter_col}' not found, skipping filter")
            
            # 통계 업데이트
            self.update_stats()
            
            self.log("=" * 50)
            self.log("✅ Processing completed successfully!")
            
            self.root.after(0, lambda: self.export_btn.config(state='normal'))
            
        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
        finally:
            self.root.after(0, lambda: self.process_btn.config(state='normal', text="▶️ Process Files"))
    
    def update_stats(self):
        if self.merged_df is None:
            return
        
        stats = f"""
Rows: {len(self.merged_df):,}
Columns: {len(self.merged_df.columns)}

Column Names:
{chr(10).join('  • ' + col for col in self.merged_df.columns[:10])}
{'  ... and more' if len(self.merged_df.columns) > 10 else ''}

Memory: {self.merged_df.memory_usage(deep=True).sum() / 1024 / 1024:.2f} MB
        """.strip()
        
        self.root.after(0, lambda: self.update_stats_text(stats))
    
    def update_stats_text(self, text):
        self.stats_text.config(state='normal')
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(1.0, text)
        self.stats_text.config(state='disabled')
    
    def export_result(self):
        if self.merged_df is None:
            messagebox.showwarning("No Data", "Process files first")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            initialfile=f"merged_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                if filename.endswith('.csv'):
                    self.merged_df.to_csv(filename, index=False, encoding='utf-8-sig')
                else:
                    self.merged_df.to_excel(filename, index=False, engine='openpyxl')
                
                self.log(f"✅ Exported to: {os.path.basename(filename)}")
                messagebox.showinfo("Success", f"File saved:\n{filename}")
                
            except Exception as e:
                self.log(f"❌ Export error: {str(e)}")
                messagebox.showerror("Export Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorTool(root)
    root.mainloop()