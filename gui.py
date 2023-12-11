from tkinter import messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd

class ExcelTemplateMaker:
    def __init__(self):
        self.root = ttk.Window(themename='darkly')
        self.root.title('Excel Template Maker')

        # Get screen width and height
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Use screen width and height for the geometry
        self.root.geometry(f'{screen_width}x{screen_height}')

        my_font = ('Avenir', 12)  # Define a common font

        # Main labelpyinstaller --noconfirm --onefile --windowed --icon
        my_label = ttk.Label(self.root, text='Make your Excel template', font=('Avenir', 20))
        my_label.grid(row=0, column=0, columnspan=2, pady=10, padx=10)

        # Labeled Frame (Left)
        labeled_frame_left = ttk.LabelFrame(self.root, text='Please fill in', labelanchor='nw', width=50, height=50)
        labeled_frame_left.grid(row=1, column=0, padx=10, pady=10, sticky='nsew')

        # Left Box
        self.left_box = ttk.Frame(labeled_frame_left)
        self.left_box.pack(anchor='nw', padx=10, pady=10)

        self.entry_values = {}  # Dictionary to store entry values

        self.create_entry_with_label('Project Name', my_font)
        self.create_entry_with_label('Experiment', my_font)
        self.create_entry_with_label('Group', my_font)
        self.create_entry_with_label('Biological individual', my_font)
        self.create_entry_with_label('Treatment/Condition', my_font)
        self.create_entry_with_label('Timepoint', my_font)
        self.create_entry_with_label('Technical Replicate', my_font)
        self.create_entry_with_label('Analytical Replicate', my_font)

        # Labeled Frame (Right)
        labeled_frame_right = ttk.LabelFrame(self.root, text='Your template', labelanchor='nw', width=50, height=50)
        labeled_frame_right.grid(row=1, column=1, padx=10, pady=10, sticky='nsew')

        # Right Box
        self.right_box = ttk.Frame(labeled_frame_right)
        self.right_box.pack(anchor='nw', padx=10, pady=10)

        # Treeview Widget for Excel-like form template
        self.columns = ['ID', 'Accession', 'Gene', 'Description', 'Your input']
        self.tree = ttk.Treeview(self.right_box, columns=self.columns, show='headings')

        for col in self.columns:
            self.tree.heading(col, text=col)

        # Set the width of the 'Your input' column to span the full width of right_box
        right_box_width = self.right_box.winfo_width()
        self.tree.column('Your input', width=390)
        self.tree.pack(expand=YES, fill=BOTH)

        # Frame for buttons at the bottom
        button_frame = ttk.Frame(self.root)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky='e')

        # Buttons in the bottom frame
        exit_button = ttk.Button(button_frame, text='Exit', bootstyle='danger', command=self.root.destroy)
        exit_button.pack(side=RIGHT, padx=10)

        generate_button = ttk.Button(button_frame, text='Generate', command=self.generate_template)
        generate_button.pack(side=RIGHT, padx=10)

        # Make sure the columns and rows expand with the window size
        for i in range(2):  # Two columns
            self.root.grid_columnconfigure(i, weight=1)

        self.root.grid_rowconfigure(1, weight=1)  # Make the second row expand vertically

        # Label to display the concatenated string
        self.concatenated_label = ttk.Label(self.right_box, text='', font=my_font)
        self.concatenated_label.pack(anchor='nw', padx=10, pady=10)

        self.root.mainloop()

    def create_entry_with_label(self, label_text, font):
        label = ttk.Label(self.left_box, text=label_text, font=font)
        label.pack(anchor='w', padx=5, pady=5)

        self.entry = ttk.Entry(self.left_box, font=font)
        self.entry.pack(anchor='w', padx=5, pady=5)

        # Set default values for the entries based on position
        if len(self.entry_values) >= 2:
            default_values = ['Ctrl', 'CT1', 'C1', 'T1', 'TR1', 'AR1']
            self.entry.insert(0, default_values[len(self.entry_values)-2])

        # Store entry value in the dictionary when changed
        self.entry.bind('<KeyRelease>', lambda event, entry=self.entry: self.update_entry_value(entry))
        self.entry_values[self.entry] = self.entry.get()

    def update_entry_value(self, entry):
        self.entry_values[entry] = entry.get()  # Update entry value in the dictionary

        # Concatenate the entry values with '_'
        concatenated_string = '_'.join(self.entry_values.values())

        # Display the concatenated string in the GUI
        self.concatenated_label.config(text=concatenated_string)

        # Update the Treeview column name dynamically
        self.update_treeview_column_name(concatenated_string)

    def update_treeview_column_name(self, new_column_name):
        self.new_column_name = new_column_name
        # Update the heading text for 'Your input' column
        self.tree.heading('Your input', text=new_column_name)

    def update_columns(self):
        # Update the 'Your input' column in the list of columns
        index = self.columns.index('Your input')
        self.columns[index] = self.new_column_name
        self.tree['columns'] = tuple(self.columns)

    def generate_template(self):
        self.update_columns()

        # Specify the file path where you want to save the Excel template
        template_file_path = 'excel_template.xlsx'

        # Create the Excel template with the updated column names
        try:
            create_excel_template(self.columns, template_file_path)
            messagebox.showinfo("Success",
                                f"Excel template with columns {self.columns} created successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")


def create_excel_template(column_names, file_path):

    # Create a DataFrame with specified column names
    template_data = pd.DataFrame(columns=column_names)

    # Save the DataFrame to an Excel file
    template_data.to_excel(file_path, index=False)


if __name__ == '__main__':
    myGUI = ExcelTemplateMaker()
