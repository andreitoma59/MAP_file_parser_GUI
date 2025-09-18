import re
import customtkinter as ctk
from tkinter import filedialog
from CTkMessagebox import CTkMessagebox
import os
import pandas as pd
from openpyxl import load_workbook

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class MapFileParserApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Map File Parser")
        self.geometry("500x400")  
        
        self.file_path = ""

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)

        self.file_path_label = ctk.CTkLabel(self, text="No file selected", anchor="w")
        self.file_path_label.grid(row=0, column=0, padx=10, pady=20, sticky="w")

        self.browse_button = ctk.CTkButton(self, text="Browse Map File", command=self.browse_file)
        self.browse_button.grid(row=0, column=1, padx=10, pady=20, sticky="e")
        
        self.selected_option = ctk.StringVar(value="python")

        self.text_file_radiobutton = ctk.CTkRadioButton(self, text="Text File", variable=self.selected_option, value="text")
        self.text_file_radiobutton.grid(row=1, column=0, columnspan=2, pady=10)

        self.python_file_radiobutton = ctk.CTkRadioButton(self, text="Python File", variable=self.selected_option, value="python")
        self.python_file_radiobutton.grid(row=2, column=0, columnspan=2, pady=10)

        self.excel_file_radiobutton = ctk.CTkRadioButton(self, text="Excel File", variable=self.selected_option, value="excel")
        self.excel_file_radiobutton.grid(row=3, column=0, columnspan=2, pady=10)
        
        self.go_button = ctk.CTkButton(self, text="Go", command=self.parse_map_file)
        self.go_button.grid(row=4, column=0, columnspan=2, pady=20)
    
    def browse_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Map files", "*.map")])
        if self.file_path:
            display_path = self.file_path
            max_length = 50
            if len(self.file_path) > max_length:
                display_path = "..." + self.file_path[-(max_length-3):]
            
            self.file_path_label.configure(text=display_path)
            CTkMessagebox(title="File Selected", message=f"Selected file: {self.file_path}", icon="info")
    
    def parse_map_file(self):
        if not self.file_path:
            CTkMessagebox(title="Error", message="Please select a map file before proceeding.", icon="cancel")
            return

        selected_format = self.selected_option.get()

        if not selected_format:
            CTkMessagebox(title="Error", message="Please select an output format.", icon="cancel")
            return

        variables_dict = self.process_map_file()
        
        if selected_format == "text":
            self.save_to_text_file(variables_dict)
        elif selected_format == "python":
            self.generate_python_file(variables_dict)
        elif selected_format == "excel":
            self.save_to_excel_file(variables_dict)

    def process_map_file(self):
        match_start = re.compile('sorted on address')
        match_stop = re.compile(r'\+\-\-\-')
        match_address = re.compile(r"^\|\s*([0-9a-fA-Fx]+)\s*\|\s*([^|]+)\s*\|")
        variables_dict = {}

        with open(self.file_path, 'r') as file:
            lines = file.readlines()
            start_value = 0
            stop_value = 0

            for i, line in enumerate(lines):
                if match_start.search(line):
                    start_value = i + 7
                    break

            for i, line in enumerate(lines[start_value:], start=start_value):
                if match_stop.search(line):
                    stop_value = i
                    break

            for i, line in enumerate(lines[start_value:stop_value], start=start_value):
                match = match_address.match(line)
                if match:
                    hexadecimal_value = match.group(1).ljust(10, '0')
                    variable_name = match.group(2).strip()
                    variables_dict[variable_name] = hexadecimal_value
        return variables_dict

    def save_to_text_file(self, variables_dict):
        output_path = os.path.splitext(self.file_path)[0] + "_variables.txt"
        with open(output_path, "w") as text_file:
            for var_name, var_value in variables_dict.items():
                text_file.write(f"{var_name}: {var_value}\n")
        CTkMessagebox(title="File Saved", message=f"Text file saved successfully: {output_path}", icon="check")

    def generate_python_file(self, variables_dict):
        class_file_content = f"""
class MapVariables:
    def __init__(self):
{chr(10).join([f"        self.{var_name} = '{var_value}'" for var_name, var_value in variables_dict.items()])}
"""

        output_path = os.path.splitext(self.file_path)[0] + "_variables.py"
        with open(output_path, "w") as class_file:
            class_file.write(class_file_content)
        
        CTkMessagebox(title="File Saved", message=f"Python file generated successfully: {output_path}", icon="check")

    def save_to_excel_file(self, variables_dict):
        output_path = os.path.splitext(self.file_path)[0] + "_variables.xlsx"
        
        df = pd.DataFrame(list(variables_dict.items()), columns=["Variable Name", "Hexadecimal Value"])
        df.to_excel(output_path, index=False)
        
        workbook = load_workbook(output_path)
        sheet = workbook.active
        
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter 
            for cell in column:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width
        
        workbook.save(output_path)
        workbook.close()
        
        CTkMessagebox(title="File Saved", message=f"Excel file saved successfully: {output_path}", icon="check")

if __name__ == "__main__":
    app = MapFileParserApp()
    app.mainloop()
