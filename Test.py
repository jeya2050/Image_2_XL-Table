import pandas as pd
from tabulate import tabulate
import tkinter as tk
import json
import tkinter.messagebox 
from tkinter import ttk, filedialog, StringVar
from tkinter import simpledialog


class ExcelDataViewer:
    
    def __init__(self, root):
        global column_frame
        # Initialize the main application window
        self.root = root
        self.root.title("Excel Data Viewer")

        file_frame = tk.LabelFrame(root, text="file upload")
        file_frame.place(height=400, width=250, x=3, y=5)
        
        column_frame=tk.LabelFrame(root,text="table_manipulation")
        column_frame.place(height=400,width=250,x=280,y=5)
        
        json_frame=tk.LabelFrame(root,text="Json Operations")
        json_frame.place(height=400,width=250,x=560,y=5)
        
        
        # Create an input frame for organizing UI elements
        self.input_frame = tk.Frame(root, width=300, height=200)
        self.input_frame.pack(side="top", padx=10, pady=10)

        # Button to upload an Excel file
        self.upload_button = tk.Button(
            file_frame, text="Upload Excel File", command=self.upload_file)
        self.upload_button.pack(pady=10)
        self.upload_button.place(x=10, y=10)

        # Dropdown for selecting a sheet
        self.sheet_variable = StringVar()
        self.sheet_variable.set("Select a sheet")
        self.sheet_dropdown = ttk.Combobox(
            file_frame, textvariable=self.sheet_variable)
        self.sheet_dropdown.pack(pady=5)
        self.sheet_dropdown.place(x=10, y=50)

        # Entry for specifying start row
        self.start_row_label = tk.Label(file_frame, text="Start Row:")
        self.start_row_label.pack()
        self.start_row_label.place(x=10, y=90)
        self.start_row_entry = tk.Entry(file_frame)
        self.start_row_entry.pack(pady=5)
        self.start_row_entry.place(x=100, y=90)

        # Entry for specifying end row
        self.end_row_label = tk.Label(file_frame, text="End Row:")
        self.end_row_label.pack()
        self.end_row_label.place(x=10, y=120)
        self.end_row_entry = tk.Entry(file_frame)
        self.end_row_entry.pack(pady=5)
        self.end_row_entry.place(x=100, y=120)

        # Button to display selected data
        self.display_button = tk.Button(
            file_frame, text="Display Data", command=self.display_data)
        self.display_button.pack(pady=5)
        self.display_button.place(x=10, y=150)

          # Reset Button
        self.reset_button = tk.Button(
            file_frame, text="Reset", command=self.reset_ui)
        self.reset_button.pack(pady=5)
        self.reset_button.place(x=100, y=150)  # Adjust position as needed
        
        # List of common datatypes
        common_datatypes = ["int", "float", "str", "bool"]  # Add more as needed

        # Dropdown for selecting a datatype
        self.datatype_variable = StringVar()
        self.datatype_variable.set("Select a datatype")
        self.datatype_dropdown = ttk.Combobox(
            column_frame, textvariable=self.datatype_variable, values=common_datatypes)
        self.datatype_dropdown.pack()
        self.datatype_dropdown.place(x=10, y=290)

        # Label for selecting a column for insertion
        self.column_selection_label = tk.Label(column_frame, text="Select Column:")
        self.column_selection_label.pack()
        self.column_selection_label.place(x=5, y=60)
        

        # # Dropdown for selecting a column
        self.column_selection = ttk.Combobox(column_frame)  # Empty for now
        self.column_selection.pack()
        self.column_selection.place(x=90, y=60)
        
        # self.column_selection["values"] = list(df.columns)
    
        # if len(list(df.columns)) > 0:
        #     self.column_selection.current(0) 

        # # Label and entry for specifying start row for insertion
        self.start_row_label1 = tk.Label(column_frame, text="Start Row:")
        self.start_row_label1.pack()
        self.start_row_label1.place(x=10, y=90)

        self.start_row_entry1 = tk.Entry(column_frame)
        self.start_row_entry1.pack()
        self.start_row_entry1.place(x=80, y=90)

        # # Label and entry for specifying end row for insertion
        self.end_row_label1 = tk.Label(column_frame, text="End Row:")
        self.end_row_label1.pack()
        self.end_row_label1.place(x=10, y=120)

        self.end_row_entry1 = tk.Entry(column_frame)
        self.end_row_entry1.pack()
        self.end_row_entry1.place(x=80, y=120)

        # Button to insert data into the specified range
        self.insert_data_button = tk.Button(
            column_frame, text="Insert Data", command=self.insert_data_range)
        self.insert_data_button.pack()
        self.insert_data_button.place(x=140, y=180)
        
        # Label and entry for entering data
        self.data_entry_label = tk.Label(column_frame, text="Enter Data:")
        self.data_entry_label.pack()
        self.data_entry_label.place(x=10, y=150)

        self.data_entry = tk.Entry(column_frame)
        self.data_entry.pack()
        self.data_entry.place(x=80, y=150)

          # Create UI elements for row deletion
        row_delete_label = tk.Label(column_frame, text="Delete Row:")
        row_delete_label.pack()
        row_delete_label.place(x=10, y=230)

        self.row_delete_entry = tk.Entry(column_frame)
        self.row_delete_entry.pack()
        self.row_delete_entry.place(x=100, y=230)

        delete_button = tk.Button(
            column_frame, text="Delete", command=self.delete_row)
        delete_button.pack()
        delete_button.place(x=170, y=250)

         # Button to add a new column to the sheet
        self.add_column_button = tk.Button(
            column_frame, text="Add Column", command=self.add_new_column)
        self.add_column_button.pack(pady=5)
        self.add_column_button.place(x=70, y=15)  # Adjust position as needed
        
        self.get_datatype_button = tk.Button(
        column_frame, text="Check Datatype", command=self.check_datatype)
        self.get_datatype_button.pack()
        self.get_datatype_button.place(x=10, y=320)
        # json_frame = tk.Frame(json_canvas)
        # json_canvas.create_window((0, 0), window=json_frame, anchor="nw")
        
        # JSON FRAME
        
        # Lists to store JSON-related information
        self.keys = []
        self.key_entry_vars = {}
        self.json_entries = []
        
        # Initialize an empty list for JSON data
        self.json_data = []
        
        # Button to add a new JSON key
        self.add_key_button = tk.Button(
            json_frame, text="Add Json", command= self.add_json)
        self.add_key_button.pack(pady=5)
        self.add_key_button.place(x=10, y=10)

        # Button to convert selected data to JSON
        self.convert_button = tk.Button(
            json_frame, text="Convert to JSON", command=self.convert_to_json)
        self.convert_button.pack(pady=5)
        self.convert_button.place(x=10, y=50)

        json_input_frame = tk.LabelFrame(root, text="Added Json")
        json_input_frame.place(height=800, width=600, x=1300, y=5)

        json_canvas = tk.Canvas(json_input_frame)
        json_canvas.pack(side="left", fill="both", expand=True)
        # json_canvas.pack(side="top", fill="both", expand=True)

        json_scrollbar = tk.Scrollbar(
            json_input_frame, orient="vertical", command=json_canvas.yview)
        json_scrollbar.pack(side="right", fill="y")
        # json_scrollbar = tk.Scrollbar(
        #     json_input_frame, orient="horizontal", command=json_canvas.xview)
        # json_scrollbar.pack(side="bottom", fill="x")
        def on_mouse_wheel(event):
            # Implement scrolling using the mouse wheel
            json_canvas.yview_scroll(-1 * (event.delta // 120), "units")
        json_canvas.configure(yscrollcommand=json_scrollbar.set)
        json_canvas.bind("<Configure>", lambda e: json_canvas.configure(scrollregion=json_canvas.bbox("all")))
        json_canvas.bind("<MouseWheel>", on_mouse_wheel)  # Bind mouse wheel event
        # Create a frame to hold the widgets inside the canvas
        self.json_content_frame = tk.Frame(json_canvas)
        json_canvas.create_window((0, 0), window=self.json_content_frame, anchor="nw")

        # Initialize the scroll region
        self.json_content_frame.update_idletasks()
        json_canvas.config(scrollregion=json_canvas.bbox("all"))
        # Button to preview JSON data
        self.preview_button = tk.Button(
            json_frame, text="Preview JSON", command=self.preview_json)
        self.preview_button.pack(pady=5)
        self.preview_button.place(x=10, y=90)  # Adjust position as needed

        
        # Button to select JSON files
        self.select_files_button = tk.Button(
            json_frame, text="Select JSON Files", command=self.select_json_files)
        self.select_files_button.pack(pady=10)
        self.select_files_button.place(x=10, y=130)

        # Button to append JSON data
        self.append_button = tk.Button(
            json_frame, text="Append JSON Data", command=self.append_json_data)
        self.append_button.pack(pady=5)
        self.append_button.place(x=10, y=170)

        # Button to preview appended JSON data
        self.preview_button = tk.Button(
            json_frame, text="Preview Appended JSON", command=self.preview_appended_json)
        self.preview_button.pack(pady=5)
        self.preview_button.place(x=10, y=210)

        # Initialize variables for storing data and widgets
        self.sheet_data = None
        self.text_widget = None
         
        self.root = root
        self.json_entries = []
        self.convert_datatype_button = tk.Button(
        column_frame, text="Convert to Datatype", command=self.convert_to_datatype)
        self.convert_datatype_button.pack()
        self.convert_datatype_button.place(x=120, y=320)
        
        # Add a button for column renaming
        self.rename_column_button = tk.Button(
            column_frame, text="Rename Column", command=self.rename_column)
        self.rename_column_button.pack()
        self.rename_column_button.place(x=10, y=180)
    
################ Add string format function....

    def data_formate(self):
        result_column=[]
        for num,i in enumerate(self.selected_rows[selected_column]):
            company=(str(i).upper()).replace(" ", "_")
            alphabetic_characters = [char for char in company if char.isalpha() or char == "_"]
            result = "".join(alphabetic_characters)
            selected_sheet = self.sheet_variable.get()
            self.sheet_data[selected_sheet].at[num,selected_column] = result
            result_column.append(result)
        # print("column",selected_column)  
        # print(result_column)
        self.show_data(self.sheet_data[selected_sheet])
        message = f"The column '{selected_column}' has formated...."
        tk.messagebox.showinfo("Result", message)
        self.format_button.destroy()
 

#############################

    def check_datatype(self):
        global selected_column
        # Get the selected column and user-selected datatype from the dropdown
        selected_column = self.column_selection.get()
        selected_datatype = self.datatype_variable.get()

        if selected_datatype == "Select a datatype":
            tk.messagebox.showinfo("No Datatype Selected", "Please select a datatype from the dropdown.")
            return

        # Check if the selected datatype matches the column's datatype, ignoring NaN values
        try:
            # Get non-NaN values from the selected column
            non_nan_values = self.selected_rows[selected_column].dropna()

            # Check if there are any non-NaN values
            if non_nan_values.empty:
                message = f"The column '{selected_column}' contains only NaN values. Selected datatype '{selected_datatype}' is not applicable."
            else:
                # Check the Python data type of the first non-NaN value
                sample_value = non_nan_values.iloc[0]
                sample_type = type(sample_value).__name__

                # Compare the selected datatype with the sample_type, case-insensitive
                if selected_datatype.lower() == sample_type.lower():
                    message = f"The column '{selected_column}' has values with the datatype '{sample_type}'.\nSelected datatype '{selected_datatype}' matches."
######################## format it button added...
                    if selected_datatype.lower() == sample_type.lower() == "str":
                        self.format_button = tk.Button(
                        column_frame, text="format it", command=self.data_formate)
                        self.format_button.pack()
                        self.format_button.place(x=40, y=350)
##########################################

                else:
                    message = f"The column '{selected_column}' has values with the datatype '{sample_type}' which does not match the selected datatype '{selected_datatype}'."

            # Display the result message
            tk.messagebox.showinfo("Datatype Match", message)
        except Exception as e:
            # If there's an error, display an error message
            message = f"Error: The column '{selected_column}' could not be converted to datatype '{selected_datatype}'.\n{str(e)}"
            tk.messagebox.showerror("Datatype Mismatch", message)

    def convert_to_datatype(self):
        # Get the selected column and user-input datatype
        selected_column = self.column_selection.get()
        user_datatype = simpledialog.askstring(
            "Input Datatype", f"Enter the datatype to convert '{selected_column}' to:")

        # Check if the user input is valid
        if user_datatype:
            try:
                # Convert the selected column to the user-input datatype, ignoring NaN values
                self.selected_rows[selected_column] = pd.to_numeric(self.selected_rows[selected_column], errors='coerce')
                self.selected_rows[selected_column] = self.selected_rows[selected_column].astype(user_datatype)
                
                # Display a success message
                message = f"The column '{selected_column}' has been converted to datatype '{user_datatype}' while ignoring NaN values."
                tk.messagebox.showinfo("Conversion Successful", message)
            except Exception as e:
                # If there's an error, display an error message
                message = f"Error: Conversion of the column '{selected_column}' to datatype '{user_datatype}' failed.\n{str(e)}"
                tk.messagebox.showerror("Conversion Error", message)


    def rename_column(self):
            # Get the selected column and user-input new column name
            selected_column = self.column_selection.get()
            new_column_name = simpledialog.askstring(
                "Rename Column", f"Enter a new name for the column '{selected_column}':")

            # Check if the user input is valid
            if new_column_name:
                try:
                    # Rename the selected column to the user-input new name
                    self.selected_rows.rename(columns={selected_column: new_column_name}, inplace=True)
                    
                    # Display a success message
                    message = f"The column '{selected_column}' has been renamed to '{new_column_name}'."
                    tk.messagebox.showinfo("Rename Successful", message)
                except Exception as e:
                    # If there's an error, display an error message
                    message = f"Error: Renaming of the column '{selected_column}' to '{new_column_name}' failed.\n{str(e)}"
                    tk.messagebox.showerror("Rename Error", message)


    

    def add_json(self):
        key = simpledialog.askstring("Column Name", "Enter column name:")
        if key:
            y_offset = len(self.json_entries) * 70

            entry_var = tk.StringVar()
            self.key_entry_vars[key] = entry_var
            label = tk.Label(self.json_content_frame, text=key.capitalize())
            label.grid(row=len(self.json_entries), column=0, padx=10, pady=10, sticky="w")

            entry = tk.Entry(self.json_content_frame, textvariable=entry_var)
            entry.grid(row=len(self.json_entries), column=1, padx=10, pady=10, sticky="w")

            selected_sheet = self.sheet_variable.get()
            if selected_sheet and self.sheet_data and selected_sheet in self.sheet_data:
                columns = list(self.sheet_data[selected_sheet].columns)
                print(list(self.sheet_data[selected_sheet].columns))
                columns_dropdown = ttk.Combobox(self.json_content_frame, values=columns)
                columns_dropdown.grid(row=len(self.json_entries), column=2, padx=10, pady=10, sticky="w")
                self.key_entry_vars[key + "_column"] = columns_dropdown
            self.json_entries.append((label, entry))
            # Update the scroll region and window position
            json_canvas = self.json_content_frame.master
            json_canvas.update_idletasks()
            json_canvas.config(scrollregion=json_canvas.bbox("all"))

    def select_json_files(self):
        # Open a dialog to select JSON files
        file_paths = filedialog.askopenfilenames(
            filetypes=[("JSON Files", "*.json")])

        # Reset previous JSON data and populate with selected files
        self.json_data = []
        for file_path in file_paths:
            with open(file_path, "r") as file:
                json_content = json.load(file)
                self.json_data.extend(json_content)

    def append_json_data(self):
        # Write appended JSON data to a file
        appended_json_file = "appended.json"
        with open(appended_json_file, "w") as file:
            json.dump(self.json_data, file, indent=4)
        print(f"Appended JSON data saved to '{appended_json_file}'")

    def preview_appended_json(self):
        # Create a preview window for appended JSON data
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Appended JSON Preview")

        # Create a text widget to display JSON data
        json_text = tk.Text(preview_window, wrap=tk.WORD, width=80, height=20)
        json_text.pack()

        # Insert and display JSON data in the text widget
        json_text.insert(tk.END, json.dumps(self.json_data, indent=4))

        # Add scrollbar for the text widget
        scrollbar = tk.Scrollbar(preview_window, command=json_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        json_text.config(yscrollcommand=scrollbar.set)
        json_text.config(xscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

    def reset_ui(self):
        # Reset various UI elements and data
        self.sheet_variable.set("Select a sheet")
        self.start_row_entry.delete(0, tk.END)
        self.end_row_entry.delete(0, tk.END)
        self.format_button.destroy()

        # Clear text widget if present
        if self.text_widget:
            self.text_widget.destroy()
        self.selected_rows = None

        # Clear added keys and their associated widgets
        self.keys = []
        for entry_var in self.key_entry_vars.values():
            entry_var.set("")
        self.key_entry_vars = {}

        # Clear existing entries from the json_entries list
        for entry in self.json_entries:
            entry.destroy()
        self.json_entries = []

        # Reset UI components to initial state
        self.sheet_dropdown["values"] = []  # Clear sheet dropdown values
        self.start_row_entry.delete(0, tk.END)  # Clear start row entry
        self.end_row_entry.delete(0, tk.END)  # Clear end row entry

        # Reset other UI components as needed

        # Reset the sheet_data attribute to None
        self.sheet_data = None
        self.text_widget = None

        # Clear text widget displaying data
        if self.text_widget:
            self.text_widget.destroy()
            self.text_widget = None

        # Update the UI
        self.root.update()

    def insert_data_range(self):
        selected_sheet = self.sheet_variable.get()
        selected_column = self.column_selection.get()
        start_row = int(self.start_row_entry1.get()
                        ) if self.start_row_entry1.get() else None
        end_row = int(self.end_row_entry1.get()
                      ) if self.end_row_entry1.get() else None
        data_to_insert = self.data_entry.get()

        # Check if all required values are provided
        if selected_sheet and self.sheet_data and selected_column and start_row and end_row and data_to_insert:
            try:
                # Insert data into specified rows and column
                for row in range(start_row, end_row + 1):
                    self.sheet_data[selected_sheet].at[row,
                                                       selected_column] = data_to_insert
                self.show_data(self.sheet_data[selected_sheet])
            except ValueError:
                print("Invalid input. Please enter valid numbers.")

    def upload_file(self):
        # Open a file dialog to select an Excel file
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")])

        # If a file was selected, display its content
        if file_path:
            self.display_xlsx_file(file_path)

    def add_new_column(self):
        # Ask the user for a new column name
        key = tk.simpledialog.askstring("Column Name", "Enter column name:")

        # If a column name is provided, add it to the sheet
        if key:
            self.add_column_to_sheet(key)

    def add_column_to_sheet(self, key):
        # Check if sheet data exists
        if self.sheet_data is not None:
            selected_sheet = self.sheet_variable.get()

            # Add a new column with empty values
            self.sheet_data[selected_sheet][key] = ""

            # Display the updated sheet data
            self.show_data(self.sheet_data[selected_sheet])

    def delete_row(self):
        # Get the row number to delete
        row_number = int(self.row_delete_entry.get())

        # Check if the row number is valid
        if 1 <= row_number:
            # Drop the selected row and display the updated data
            self.selected_rows = self.selected_rows.drop(row_number)
            self.show_data(self.selected_rows)
        else:
            print("Invalid row number. Please enter a valid row number.")

    def display_xlsx_file(self, file_path):
        # Read all sheets in the Excel file
        self.sheet_data = pd.read_excel(file_path, sheet_name=None)

        # Update sheet dropdown values with sheet names
        self.sheet_dropdown["values"] = list(self.sheet_data.keys())

    def display_data(self):
        selected_sheet = self.sheet_variable.get()
        start_row = int(self.start_row_entry.get()
                        ) if self.start_row_entry.get() else None
        end_row = int(self.end_row_entry.get()
                      ) if self.end_row_entry.get() else None

        if selected_sheet and self.sheet_data:
            # Get selected sheet data and rows based on input
            selected_df = self.sheet_data[selected_sheet]
            # Get the column names from the DataFrame
            column_names = selected_df.columns.tolist()
             # Update the column selection dropdown with column names
            self.column_selection['values'] = column_names
            self.column_selection.set(column_names[0])  # Set the default selection
            
            if start_row is not None and end_row is not None:
                selected_df = selected_df.iloc[start_row:end_row + 1]

            # Display the selected data
            self.show_data(selected_df)

            # Store the selected rows for JSON conversion
        self.selected_rows = selected_df

        
    def show_data(self, df):
        # Destroy the previous text widget, if it exists
        if self.text_widget:
            self.text_widget.destroy()

        # Create a new text widget for displaying data
        self.text_widget = tk.Text(
            self.root, wrap="none", width=150, height=30)
        self.text_widget.pack(fill="none", expand=True)
        self.text_widget.place(x=10, y=450)

        # Add row numbers as the first column
        df_with_row_numbers = df.reset_index()
        
        df_with_row_numbers.index += 1  # Start index from 1

        # Format large numbers to avoid scientific notation
        def float_format(x): return '{:,.2f}'.format(
            x) if isinstance(x, (float, int)) else x
        formatted_df = df_with_row_numbers.applymap(float_format)

        # Create a table representation of the formatted DataFrame
        table = tabulate(formatted_df, headers='keys',
                         tablefmt='grid', showindex=False)
        self.text_widget.insert("1.0", table)

    def convert_to_json(self):
        selected_sheet = self.sheet_variable.get()
        if selected_sheet and self.sheet_data and hasattr(self, 'selected_rows'):
            # Get the desired JSON file name from the user through a dialog
            json_file_name = self.get_json_file_name()
            if json_file_name:
                # Create JSON data from selected rows and keys
                json_data = self.create_json(self.selected_rows)

                # Write JSON data to a file
                with open(json_file_name, "w") as f:
                    json.dump(json_data, f, indent=4)
                print(f"JSON data saved to '{json_file_name}'")

    def get_json_file_name(self):
        # Show a dialog to get the desired JSON file name from the user
        root = tk.Tk()
        root.withdraw()
        json_file_name = simpledialog.askstring(
            "JSON File Name", "Enter the desired JSON file name:")
        root.destroy()

        # Ensure the file name ends with ".json"
        if json_file_name and not json_file_name.endswith(".json"):
            json_file_name += ".json"
        return json_file_name

    def create_json(self, df):
        json_data = []
        for _, row in df.iterrows():
            data = {}
            for key, value in self.key_entry_vars.items():
                if key.endswith("_column"):  # Check if it's a column mapping entry
                    column_name = value.get()
                    # Remove '_column' suffix
                    if column_name:
                        # Check if column_name is not empty
                        json_key = key.replace("_column", "")
                        data[json_key] = row[column_name]
                else:
                    data[key] = value.get()
            json_data.append(data)
        return json_data

    def preview_json(self):
        selected_sheet = self.sheet_variable.get()
        if selected_sheet and self.sheet_data and hasattr(self, 'selected_rows'):
            # Create JSON data from selected rows and keys
            json_data = self.create_json(self.selected_rows)
            # Show a popup with the JSON preview
            self.show_json_preview(json_data)

    def show_json_preview(self, json_data):
        popup = tk.Toplevel(self.root)
        popup.title("JSON Preview")

        json_text = tk.Text(popup, wrap=tk.WORD, width=80, height=20)
        json_text.pack()

        json_text.insert(tk.END, json.dumps(json_data, indent=4))

        scrollbar = tk.Scrollbar(popup, command=json_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        json_text.config(yscrollcommand=scrollbar.set)

        close_button = tk.Button(popup, text="Close", command=popup.destroy)
        close_button.pack()

#  initializes the GUI application and starts the main event loop.

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelDataViewer(root)
    root.mainloop()