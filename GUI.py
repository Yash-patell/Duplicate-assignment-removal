from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd
import os 

file_path = None
new_file_path = None


def open_file_dialog():
    global file_path
    file_path = filedialog.askopenfilename(title="Select a file",
                                           filetypes=[("All Files", "*.*"),
                                                      ("Excel File", "*.xlsx"),
                                                     ])
    if file_path:
        file_path_label.configure(text=file_path)


# Warning box
def show_warning_dialog(file_path, new_file_path):
    file_name = os.path.basename(file_path)
    new_file_name = os.path.basename(new_file_path)
    response = messagebox.askyesnocancel(message=f"Warning The file {file_name} already exists.\n"
                                           f"Do you want to overwrite it or save as {new_file_path}?",
                                   icon="question", title="File already exists")
    return response 

# Function to save the DataFrame to a unique file name
def save_with_unique_name(df_unique, base_path):
    counter = 1
    new_file_path = f"{base_path}_{counter}.xlsx"

    while os.path.exists(new_file_path):
        counter += 1
        new_file_path = f"{base_path.rstrip('.xlsx')}_{counter}.xlsx"

    df_unique.to_excel(new_file_path, index=False)
    print(f"Data saved to {new_file_path}")

    
def proceed():
    global file_path, new_file_path
    
    if file_path:
        df = pd.read_excel(file_path)
        # Convert submission_time to datetime
        df['submission_date'] = pd.to_datetime(df['submission_date'])

        # Identify duplicate submissions
        duplicates = df[df.duplicated(['Roll_no', 'assignment'], keep=False)]

        # Generate a list of students with duplicate submissions
        duplicate_students = duplicates['Student_name'].unique()

        # Remove duplicate submissions, keeping only the first submission
        df_unique = df.drop_duplicates(['Roll_no', 'assignment'], keep='first')

        # Define the output file path
        output_file_path = f"{os.path.splitext(file_path)[0]}_cleaned.xlsx"

        if os.path.exists(output_file_path):
            new_file_path = f"{os.path.splitext(file_path)[0]}_cleaned_new.xlsx"
            response = show_warning_dialog(output_file_path, new_file_path)
            messagebox.showinfo("Success","Duplicates Removed and Excel file saved")

            if response is None:
                return  # Cancelled, do nothing
            
            elif response:
                # Save with a unique name
                save_with_unique_name(df_unique, output_file_path)

            else:
                # User wants to overwrite
                df_unique.to_excel(output_file_path, index=False)
                print(f"Data saved to {output_file_path}")

        else:
            # Save the cleaned data to the specified file
            df_unique.to_excel(output_file_path, index=False)
            print(f"Data saved to {output_file_path}")

    else:
        messagebox.showwarning("No File Selected", "Please select an Excel file.")



# UI
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("dark-blue")

# App frame
app = ctk.CTk()
app.geometry("700x520")
app.title("Duplicate submission removal")

# UI
CF = ("Helvetica", 18)
title = ctk.CTkLabel(app, text="Insert Excel file", font=CF)
title.pack(padx=(10,0), pady=(50,10))

# Create a button to open the file dialog
insert_file_button = ctk.CTkButton(app, text="Insert File", font=("Times New Roman", 18),text_color="black",
command=open_file_dialog,fg_color='white',hover_color="grey",border_color='black')
insert_file_button.pack(padx=(0,0), pady=(100,0))

# Display selected file path
file_path_label = ctk.CTkLabel(app, text="No file selected", font=("Helvetica", 14))
file_path_label.pack(padx=(0,0), pady=(20,0))

# Create a frame to hold the buttons
button_frame = ctk.CTkFrame(app)
button_frame.pack(padx=(500,0), pady=(210,0))

# Create the "Start" button
start_button = ctk.CTkButton(button_frame, text="Start",font=("Times New Roman", 14),text_color="black", command=proceed,width=80,height=35,fg_color='white',hover_color='#AEE398',border_color='black')
start_button.pack(side="left", padx=(0,10), pady=(0,0))

# Create the "Quit" button
quit_button = ctk.CTkButton(button_frame, text="Quit", font=("Times New Roman", 14),text_color_disabled='black',text_color='black', command=app.quit,width=50,height=35,fg_color='white',hover_color='#DD4C2D',border_color='black')
quit_button.pack(side="right", padx=(10,5), pady=(0,0))




app.mainloop()
