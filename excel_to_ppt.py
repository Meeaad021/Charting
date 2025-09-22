import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

# Input/Output files (defaults)
excel_file = "data.xlsx"
template_ppt = "PPT Master Template.pptx"
output_ppt = "excel_barcharts.pptx"

BAR_COLOUR = RGBColor(0x00, 0xBC, 0xF2)

# Available chart types
CHART_TYPES = {
    "Bar Chart": XL_CHART_TYPE.BAR_CLUSTERED,
    "Column Chart": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "Pie Chart": XL_CHART_TYPE.PIE,
    "Line Chart": XL_CHART_TYPE.LINE,
    "Area Chart": XL_CHART_TYPE.AREA,
    "Doughnut Chart": XL_CHART_TYPE.DOUGHNUT,
    "Stacked Bar": XL_CHART_TYPE.BAR_STACKED,
    "Stacked Column": XL_CHART_TYPE.COLUMN_STACKED,
}

class ChartConfigUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint Chart Generator - Dynamic Configuration")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        
        # Variables
        self.excel_path = tk.StringVar(value=excel_file)
        self.template_path = tk.StringVar(value=template_ppt)
        self.output_path = tk.StringVar(value=output_ppt)
        self.starting_slide = tk.IntVar(value=3)
        
        # Chart selections and sheet info
        self.chart_selections = {}
        self.sheet_enabled = {}
        self.valid_sheets = []
        self.all_sheets_info = []
        
        self.setup_ui()
        self.load_excel_info()
        
    def setup_ui(self):
        # Main frame with scrollbar
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack scrollable elements
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Main content frame
        main_frame = ttk.Frame(self.scrollable_frame, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # Title
        title_label = ttk.Label(main_frame, text="PowerPoint Chart Generator - Dynamic Edition", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=row, column=0, columnspan=3, pady=(0, 20))
        row += 1
        
        # File selection section
        files_frame = ttk.LabelFrame(main_frame, text="File Configuration", padding="10")
        files_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        files_frame.columnconfigure(1, weight=1)
        row += 1
        
        # Excel file
        ttk.Label(files_frame, text="Excel File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        ttk.Entry(files_frame, textvariable=self.excel_path, width=60).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(files_frame, text="Browse", command=self.browse_excel).grid(row=0, column=2)
        
        # Template file
        ttk.Label(files_frame, text="Template PPT:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        ttk.Entry(files_frame, textvariable=self.template_path, width=60).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 5), pady=(5, 0))
        ttk.Button(files_frame, text="Browse", command=self.browse_template).grid(row=1, column=2, pady=(5, 0))
        
        # Output file
        ttk.Label(files_frame, text="Output PPT:").grid(row=2, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        ttk.Entry(files_frame, textvariable=self.output_path, width=60).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(0, 5), pady=(5, 0))
        ttk.Button(files_frame, text="Save As", command=self.browse_output).grid(row=2, column=2, pady=(5, 0))
        
        # Starting slide configuration
        config_frame = ttk.Frame(files_frame)
        config_frame.grid(row=3, column=0, columnspan=3, pady=(10, 0), sticky=tk.W)
        
        ttk.Label(config_frame, text="Starting Slide Number:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Spinbox(config_frame, from_=1, to=100, textvariable=self.starting_slide, width=5).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(config_frame, text="🔄 Refresh Excel Data", command=self.load_excel_info).pack(side=tk.LEFT, padx=10)
        
        # Info section
        self.info_frame = ttk.LabelFrame(main_frame, text="Excel Data Analysis", padding="10")
        self.info_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        row += 1
        
        self.info_label = ttk.Label(self.info_frame, text="Load Excel file to analyze available data...")
        self.info_label.grid(row=0, column=0, sticky=tk.W)
        
        # Chart selection section
        self.chart_frame = ttk.LabelFrame(main_frame, text="Chart Configuration for Each Sheet", padding="10")
        self.chart_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self.chart_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(row, weight=1)
        row += 1
        
        # Quick selection section
        quick_frame = ttk.LabelFrame(main_frame, text="Batch Operations", padding="10")
        quick_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        row += 1
        
        # Quick selection buttons row 1
        quick_row1 = ttk.Frame(quick_frame)
        quick_row1.pack(fill=tk.X, pady=2)
        
        ttk.Button(quick_row1, text="✓ Enable All Sheets", command=self.enable_all_sheets).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_row1, text="✗ Disable All Sheets", command=self.disable_all_sheets).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_row1, text="📊 All Bar Charts", command=lambda: self.set_all_charts("Bar Chart")).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_row1, text="📈 All Column Charts", command=lambda: self.set_all_charts("Column Chart")).pack(side=tk.LEFT, padx=2)
        
        # Quick selection buttons row 2
        quick_row2 = ttk.Frame(quick_frame)
        quick_row2.pack(fill=tk.X, pady=2)
        
        ttk.Button(quick_row2, text="🥧 All Pie Charts", command=lambda: self.set_all_charts("Pie Chart")).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_row2, text="📉 All Line Charts", command=lambda: self.set_all_charts("Line Chart")).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_row2, text="🎯 Mixed Pattern", command=self.set_mixed_charts).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_row2, text="🎲 Random Types", command=self.set_random_charts).pack(side=tk.LEFT, padx=2)
        
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=3, pady=(20, 0))
        
        ttk.Button(button_frame, text="👁️ Preview Configuration", command=self.preview_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="🚀 Generate PowerPoint", command=self.generate_ppt, 
                  style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="❌ Exit", command=self.root.quit).pack(side=tk.LEFT, padx=5)
        
        # Bind mousewheel to canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
    def browse_excel(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_path.set(filename)
            self.load_excel_info()
    
    def browse_template(self):
        filename = filedialog.askopenfilename(
            title="Select PowerPoint Template",
            filetypes=[("PowerPoint files", "*.pptx *.ppt"), ("All files", "*.*")]
        )
        if filename:
            self.template_path.set(filename)
    
    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="Save PowerPoint As",
            filetypes=[("PowerPoint files", "*.pptx")],
            defaultextension=".pptx"
        )
        if filename:
            self.output_path.set(filename)
    
    def load_excel_info(self):
        try:
            if not os.path.exists(self.excel_path.get()):
                self.info_label.config(text="❌ Excel file not found")
                return
            
            # Load all Excel sheets
            sheets = pd.read_excel(self.excel_path.get(), sheet_name=None)
            
            # Analyze all sheets
            self.all_sheets_info = []
            self.valid_sheets = []
            
            for sheet_name, df in sheets.items():
                sheet_info = {
                    'name': sheet_name,
                    'total_rows': len(df),
                    'total_columns': len(df.columns),
                    'is_valid': False,
                    'valid_rows': 0,
                    'has_numeric_data': False,
                    'column_names': list(df.columns) if not df.empty else []
                }
                
                if not df.empty and len(df.columns) >= 2:
                    # Test data cleaning process
                    test_df = df.iloc[:, :2].dropna()
                    test_df = test_df[~test_df.iloc[:, 0].astype(str).str.startswith("Base")]
                    
                    # Try to convert second column to numeric
                    numeric_col = pd.to_numeric(test_df.iloc[:, 1], errors='coerce')
                    test_df = test_df[numeric_col.notna()]
                    
                    if not test_df.empty:
                        sheet_info['is_valid'] = True
                        sheet_info['valid_rows'] = len(test_df)
                        sheet_info['has_numeric_data'] = True
                        self.valid_sheets.append(sheet_name)
                
                self.all_sheets_info.append(sheet_info)
            
            # Update info display
            total_sheets = len(sheets)
            valid_sheets_count = len(self.valid_sheets)
            
            info_text = f"📊 Excel Analysis Results:\n"
            info_text += f"   • Total sheets found: {total_sheets}\n"
            info_text += f"   • Sheets with valid chart data: {valid_sheets_count}\n"
            info_text += f"   • Charts will start from slide: {self.starting_slide.get()}\n"
            
            if valid_sheets_count > 0:
                info_text += f"   • Valid sheets: {', '.join(self.valid_sheets[:5])}"
                if len(self.valid_sheets) > 5:
                    info_text += f" and {len(self.valid_sheets)-5} more..."
            
            self.info_label.config(text=info_text)
            
            # Create dynamic chart selectors
            self.create_dynamic_selectors()
            
        except Exception as e:
            self.info_label.config(text=f"❌ Error analyzing Excel file: {str(e)}")
    
    def create_dynamic_selectors(self):
        # Clear existing selectors
        for widget in self.chart_frame.winfo_children():
            widget.destroy()
        
        self.chart_selections.clear()
        self.sheet_enabled.clear()
        
        if not self.all_sheets_info:
            ttk.Label(self.chart_frame, text="No sheets found in Excel file").pack(pady=20)
            return
        
        # Create scrollable frame for selectors
        selector_canvas = tk.Canvas(self.chart_frame, height=400)
        selector_scrollbar = ttk.Scrollbar(self.chart_frame, orient="vertical", command=selector_canvas.yview)
        selector_scrollable_frame = ttk.Frame(selector_canvas)
        
        selector_scrollable_frame.bind(
            "<Configure>",
            lambda e: selector_canvas.configure(scrollregion=selector_canvas.bbox("all"))
        )
        
        selector_canvas.create_window((0, 0), window=selector_scrollable_frame, anchor="nw")
        selector_canvas.configure(yscrollcommand=selector_scrollbar.set)
        
        # Header
        header_frame = ttk.Frame(selector_scrollable_frame)
        header_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(header_frame, text="Include", font=('Arial', 9, 'bold'), width=8).grid(row=0, column=0)
        ttk.Label(header_frame, text="Sheet Name", font=('Arial', 9, 'bold'), width=20).grid(row=0, column=1, sticky=tk.W, padx=5)
        ttk.Label(header_frame, text="Data Info", font=('Arial', 9, 'bold'), width=15).grid(row=0, column=2)
        ttk.Label(header_frame, text="Chart Type", font=('Arial', 9, 'bold'), width=18).grid(row=0, column=3)
        ttk.Label(header_frame, text="Slide #", font=('Arial', 9, 'bold'), width=8).grid(row=0, column=4)
        
        # Create selector for each sheet
        enabled_count = 0
        for i, sheet_info in enumerate(self.all_sheets_info):
            sheet_name = sheet_info['name']
            
            frame = ttk.Frame(selector_scrollable_frame)
            frame.pack(fill=tk.X, padx=5, pady=1)
            
            # Enable/Disable checkbox
            enabled_var = tk.BooleanVar(value=sheet_info['is_valid'])
            self.sheet_enabled[sheet_name] = enabled_var
            
            checkbox = ttk.Checkbutton(frame, variable=enabled_var, command=self.update_slide_numbers)
            checkbox.grid(row=0, column=0, padx=5)
            
            # Sheet name with status indicator
            name_text = sheet_name
            if sheet_info['is_valid']:
                name_text = f"✅ {sheet_name}"
            else:
                name_text = f"❌ {sheet_name}"
                enabled_var.set(False)  # Disable invalid sheets
            
            name_label = ttk.Label(frame, text=name_text, width=22)
            name_label.grid(row=0, column=1, sticky=tk.W, padx=5)
            
            # Data info
            if sheet_info['is_valid']:
                data_info = f"{sheet_info['valid_rows']} rows"
            else:
                data_info = f"No valid data"
            
            ttk.Label(frame, text=data_info, width=15).grid(row=0, column=2)
            
            # Chart type selector
            chart_var = tk.StringVar(value="Bar Chart")
            self.chart_selections[sheet_name] = chart_var
            
            if sheet_info['is_valid']:
                combo = ttk.Combobox(frame, textvariable=chart_var, values=list(CHART_TYPES.keys()), 
                                   state="readonly", width=16)
            else:
                combo = ttk.Combobox(frame, textvariable=chart_var, values=list(CHART_TYPES.keys()), 
                                   state="disabled", width=16)
            
            combo.grid(row=0, column=3, padx=5)
            
            # Slide number (will be updated dynamically)
            slide_label = ttk.Label(frame, text="", width=8)
            slide_label.grid(row=0, column=4)
            
            # Store reference to slide label for updates
            frame.slide_label = slide_label
            frame.sheet_name = sheet_name
            
            if sheet_info['is_valid'] and enabled_var.get():
                enabled_count += 1
        
        # Pack the canvas and scrollbar
        selector_canvas.pack(side="left", fill="both", expand=True)
        selector_scrollbar.pack(side="right", fill="y")
        
        # Update slide numbers
        self.update_slide_numbers()
        
        # Bind mousewheel to selector canvas
        def _on_selector_mousewheel(event):
            selector_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        selector_canvas.bind_all("<Button-4>", lambda e: selector_canvas.yview_scroll(-1, "units"))
        selector_canvas.bind_all("<Button-5>", lambda e: selector_canvas.yview_scroll(1, "units"))
    
    def update_slide_numbers(self):
        """Update slide numbers based on enabled sheets"""
        slide_num = self.starting_slide.get()
        
        # Find all frames with slide labels
        for widget in self.chart_frame.winfo_children():
            if isinstance(widget, tk.Canvas):
                scrollable_frame = widget.winfo_children()[0] if widget.winfo_children() else None
                if scrollable_frame:
                    for frame in scrollable_frame.winfo_children():
                        if hasattr(frame, 'slide_label') and hasattr(frame, 'sheet_name'):
                            if self.sheet_enabled.get(frame.sheet_name, tk.BooleanVar()).get():
                                frame.slide_label.config(text=f"Slide {slide_num}")
                                slide_num += 1
                            else:
                                frame.slide_label.config(text="Disabled")
    
    def enable_all_sheets(self):
        for sheet_name, var in self.sheet_enabled.items():
            # Only enable sheets that have valid data
            sheet_info = next((s for s in self.all_sheets_info if s['name'] == sheet_name), None)
            if sheet_info and sheet_info['is_valid']:
                var.set(True)
        self.update_slide_numbers()
    
    def disable_all_sheets(self):
        for var in self.sheet_enabled.values():
            var.set(False)
        self.update_slide_numbers()
    
    def set_all_charts(self, chart_type):
        for sheet_name, var in self.chart_selections.items():
            if self.sheet_enabled.get(sheet_name, tk.BooleanVar()).get():
                var.set(chart_type)
    
    def set_mixed_charts(self):
        chart_types = ["Bar Chart", "Column Chart", "Pie Chart", "Line Chart"]
        enabled_sheets = [name for name, var in self.sheet_enabled.items() if var.get()]
        
        for i, sheet_name in enumerate(enabled_sheets):
            if sheet_name in self.chart_selections:
                self.chart_selections[sheet_name].set(chart_types[i % len(chart_types)])
    
    def set_random_charts(self):
        import random
        chart_types = list(CHART_TYPES.keys())
        
        for sheet_name, var in self.chart_selections.items():
            if self.sheet_enabled.get(sheet_name, tk.BooleanVar()).get():
                var.set(random.choice(chart_types))
    
    def get_enabled_sheets(self):
        """Get list of enabled sheets with their configuration"""
        enabled_sheets = []
        slide_num = self.starting_slide.get()
        
        for sheet_info in self.all_sheets_info:
            sheet_name = sheet_info['name']
            if self.sheet_enabled.get(sheet_name, tk.BooleanVar()).get() and sheet_info['is_valid']:
                enabled_sheets.append({
                    'name': sheet_name,
                    'chart_type': self.chart_selections[sheet_name].get(),
                    'slide_number': slide_num,
                    'data_rows': sheet_info['valid_rows']
                })
                slide_num += 1
        
        return enabled_sheets
    
    def preview_settings(self):
        enabled_sheets = self.get_enabled_sheets()
        
        if not enabled_sheets:
            messagebox.showwarning("No Sheets Selected", "Please enable at least one sheet with valid data!")
            return
        
        preview_text = "📋 POWERPOINT GENERATION PREVIEW\n" + "="*60 + "\n\n"
        preview_text += f"📁 Excel File: {os.path.basename(self.excel_path.get())}\n"
        preview_text += f"📄 Template: {os.path.basename(self.template_path.get())}\n"
        preview_text += f"💾 Output: {os.path.basename(self.output_path.get())}\n\n"
        
        preview_text += f"🎯 Starting Slide: {self.starting_slide.get()}\n"
        preview_text += f"📊 Total Charts to Create: {len(enabled_sheets)}\n\n"
        
        preview_text += "📈 CHART CONFIGURATION:\n" + "-"*40 + "\n"
        
        for sheet_config in enabled_sheets:
            preview_text += f"Slide {sheet_config['slide_number']:2d}: {sheet_config['name']:<25} → "
            preview_text += f"{sheet_config['chart_type']:<15} ({sheet_config['data_rows']} data points)\n"
        
        final_slide = enabled_sheets[-1]['slide_number'] if enabled_sheets else self.starting_slide.get()
        template_slides = 2  # Assuming template has 2 slides
        preview_text += f"\n🎯 Final presentation will have {max(template_slides, final_slide)} slides"
        
        # Show preview dialog
        self.show_preview_dialog(preview_text)
    
    def show_preview_dialog(self, preview_text):
        dialog = tk.Toplevel(self.root)
        dialog.title("Preview Configuration")
        dialog.geometry("800x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Text widget with scrollbar
        text_frame = ttk.Frame(dialog)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=('Courier', 10))
        text_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=text_scrollbar.set)
        
        text_widget.pack(side="left", fill="both", expand=True)
        text_scrollbar.pack(side="right", fill="y")
        
        text_widget.insert(tk.END, preview_text)
        text_widget.config(state=tk.DISABLED)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Generate Now", 
                  command=lambda: [dialog.destroy(), self.generate_ppt()]).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Close", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def generate_ppt(self):
        enabled_sheets = self.get_enabled_sheets()
        
        if not enabled_sheets:
            messagebox.showerror("Error", "No sheets selected for chart generation!")
            return
        
        try:
            # Show progress window
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Generating PowerPoint...")
            progress_window.geometry("500x200")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            progress_label = ttk.Label(progress_window, text="Initializing...", font=('Arial', 10))
            progress_label.pack(pady=20)
            
            progress_bar = ttk.Progressbar(progress_window, mode='determinate', length=400)
            progress_bar.pack(pady=10)
            progress_bar['maximum'] = len(enabled_sheets) + 2
            
            detail_label = ttk.Label(progress_window, text="", font=('Arial', 9), foreground='gray')
            detail_label.pack(pady=5)
            
            self.root.update()
            
            # Generate the PowerPoint
            self.create_powerpoint(progress_label, progress_bar, detail_label, enabled_sheets)
            
            progress_window.destroy()
            
            success_msg = f"PowerPoint created successfully!\n\n"
            success_msg += f"📊 Charts created: {len(enabled_sheets)}\n"
            success_msg += f"💾 Saved as: {os.path.basename(self.output_path.get())}\n\n"
            success_msg += f"📂 Full path: {self.output_path.get()}"
            
            messagebox.showinfo("Success! 🎉", success_msg)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create PowerPoint:\n\n{str(e)}")
    
    def create_powerpoint(self, progress_label, progress_bar, detail_label, enabled_sheets):
        # Load files
        progress_label.config(text="Loading Excel data...")
        detail_label.config(text=f"Reading {os.path.basename(self.excel_path.get())}")
        self.root.update()
        
        sheets = pd.read_excel(self.excel_path.get(), sheet_name=None)
        progress_bar['value'] += 1
        self.root.update()
        
        progress_label.config(text="Loading PowerPoint template...")
        detail_label.config(text=f"Opening {os.path.basename(self.template_path.get())}")
        self.root.update()
        
        prs = Presentation(self.template_path.get())
        slide_layout = prs.slide_layouts[min(2, len(prs.slide_layouts)-1)]
        progress_bar['value'] += 1
        self.root.update()
        
        # Process each enabled sheet
        for i, sheet_config in enumerate(enabled_sheets):
            sheet_name = sheet_config['name']
            chart_type_name = sheet_config['chart_type']
            chart_type = CHART_TYPES[chart_type_name]
            
            progress_label.config(text=f"Creating chart {i+1} of {len(enabled_sheets)}")
            detail_label.config(text=f"Processing {sheet_name} → {chart_type_name}")
            self.root.update()
            
            # Get and clean data
            df = sheets[sheet_name]
            df = df.iloc[:, :2].dropna()
            df = df[~df.iloc[:, 0].astype(str).str.startswith("Base")]
            df.iloc[:, 1] = pd.to_numeric(df.iloc[:, 1], errors="coerce")
            df = df.dropna()
            
            # Create slide
            slide = prs.slides.add_slide(slide_layout)
            
            # Add title
            title_text = f"{chart_type_name} - {sheet_name}"
            if slide.shapes.title:
                slide.shapes.title.text = title_text
            else:
                title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
                title_frame = title_box.text_frame
                p = title_frame.paragraphs[0]
                p.text = title_text
                p.font.size = Pt(24)
            
            # Create chart
            chart_data = CategoryChartData()
            chart_data.categories = df.iloc[:, 0].astype(str)
            chart_data.add_series("Values", df.iloc[:, 1].astype(float))
            
            x, y, cx, cy = Inches(1), Inches(2), Inches(8), Inches(5)
            chart_shape = slide.shapes.add_chart(chart_type, x, y, cx, cy, chart_data)
            chart = chart_shape.chart
            
            # Format chart
            self.format_chart(chart, chart_type)
            
            progress_bar['value'] += 1
            self.root.update()
        
        # Save presentation
        progress_label.config(text="Saving PowerPoint presentation...")
        detail_label.config(text=f"Writing to {os.path.basename(self.output_path.get())}")
        self.root.update()
        
        prs.save(self.output_path.get())
    
    def format_chart(self, chart, chart_type):
        """Apply formatting based on chart type"""
        try:
            if chart_type in [XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT]:
                # Pie/Doughnut chart formatting
                chart.has_legend = True
                chart.legend.position = XL_LEGEND_POSITION.RIGHT
                chart.plots[0].has_data_labels = True
                chart.plots[0].data_labels.show_percentage = True
                chart.plots[0].data_labels.show_category_name = False
                chart.plots[0].data_labels.show_value = True
            else:
                # Bar, Column, Line, Area charts
                chart.has_legend = False
                
                # Only set axis properties for charts that have axes
                if hasattr(chart, 'value_axis') and hasattr(chart, 'category_axis'):
                    try:
                        chart.value_axis.has_major_gridlines = False
                        chart.value_axis.has_minor_gridlines = False
                        chart.category_axis.has_major_gridlines = False
                        chart.category_axis.has_minor_gridlines = False
                        chart.value_axis.tick_labels.font.size = Pt(10)
                        chart.category_axis.tick_labels.font.size = Pt(10)
                    except:
                        pass  # Some chart types might not support all axis properties
                
                # Add data labels
                try:
                    chart.plots[0].has_data_labels = True
                    chart.plots[0].data_labels.show_value = True
                except:
                    pass  # Some chart types might not support data labels
            
            # Set colors for all series
            for series in chart.series:
                try:
                    if chart_type in [XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT]:
                        # For pie charts, color individual points
                        for point in series.points:
                            point.format.fill.solid()
                            point.format.fill.fore_color.rgb = BAR_COLOUR
                    else:
                        # For other chart types
                        series.format.fill.solid()
                        series.format.fill.fore_color.rgb = BAR_COLOUR
                        
                        # Set gap width for bar/column charts
                        if hasattr(series, 'gap_width') and chart_type in [
                            XL_CHART_TYPE.BAR_CLUSTERED, XL_CHART_TYPE.COLUMN_CLUSTERED,
                            XL_CHART_TYPE.BAR_STACKED, XL_CHART_TYPE.COLUMN_STACKED
                        ]:
                            series.gap_width = 50
                except:
                    pass  # Some formatting might not be supported for all chart types
            
            # Set data label font size
            try:
                chart.plots[0].data_labels.font.size = Pt(9)
            except:
                pass
                
        except Exception as e:
            print(f"Warning: Could not apply all formatting to chart: {e}")

def main():
    root = tk.Tk()
    app = ChartConfigUI(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()