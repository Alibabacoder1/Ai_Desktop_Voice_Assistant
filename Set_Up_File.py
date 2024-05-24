import tkinter as tk
import sqlite3
import os
import subprocess
import pythoncom
import pywintypes
from win32com.client import Dispatch

popup_window = None

# Create the database and 'user' table if they don't exist
conn = sqlite3.connect('Desktop_Ai_Assistant.db')
cursor = conn.cursor()

cursor.execute('''CREATE TABLE IF NOT EXISTS user (
                    user TEXT PRIMARY KEY,
                    data TEXT
                )''')

conn.commit()
conn.close()

# Create the main window
root = tk.Tk()
root.title("User Information and Folder Paths")

# Set the window size to 15cm x 15cm (approximately)
root.geometry("620x520")  # 1 cm is approximately 40 pixels

# Create and arrange the input fields and labels
container = tk.Frame(root, bg="#b7f1ea")
container.pack(padx=20, pady=20, fill="both", expand=True)  # Fill both horizontally and vertically

# Add the title label
title_label = tk.Label(container, text="User Information And Folder Paths", font=("Times New Roman", 24), bg="#b7f1ea")
title_label.grid(row=0, column=0, columnspan=2, pady=20)

# Add the name label and entry field
name_label = tk.Label(container, text="Name:", font=("Times New Roman", 17), bg="#b7f1ea")
name_label.grid(row=1, column=0, sticky="w")
name_entry = tk.Entry(container, font=("Times New Roman", 14), width=30)  # Increased width
name_entry.grid(row=1, column=1, padx=10, pady=10)

# Add the email label and entry field
email_label = tk.Label(container, text="Email:", font=("Times New Roman", 17), bg="#b7f1ea")
email_label.grid(row=2, column=0, sticky="w")
email_entry = tk.Entry(container, font=("Times New Roman", 14), width=30)  # Increased width
email_entry.grid(row=2, column=1, padx=10, pady=10)

# Add the phone number label and entry field
phone_label = tk.Label(container, text="Phone Number:", font=("Times New Roman", 17), bg="#b7f1ea")
phone_label.grid(row=3, column=0, sticky="w")
phone_entry = tk.Entry(container, font=("Times New Roman", 14), width=30)  # Increased width
phone_entry.grid(row=3, column=1, padx=10, pady=10)

# Add the city name label and entry field
city_label = tk.Label(container, text="City Name:", font=("Times New Roman", 17), bg="#b7f1ea")
city_label.grid(row=4, column=0, sticky="w")
city_entry = tk.Entry(container, font=("Times New Roman", 14), width=30)  # Increased width
city_entry.grid(row=4, column=1, padx=10, pady=10)

# Add the folder path of music
music_label = tk.Label(container, text="Folder Path for Music:", font=("Times New Roman", 17), bg="#b7f1ea")
music_label.grid(row=5, column=0, sticky="w")
music_entry = tk.Entry(container, font=("Times New Roman", 14), width=30)  # Increased width
music_entry.grid(row=5, column=1, padx=10, pady=10)

# Add the folder path of screenshot
screenshot_label = tk.Label(container, text="Folder Path for Screenshot :", font=("Times New Roman", 17), bg="#b7f1ea")
screenshot_label.grid(row=6, column=0, sticky="w")
screenshot_entry = tk.Entry(container, font=("Times New Roman", 14), width=30)  # Increased width
screenshot_entry.grid(row=6, column=1, padx=10, pady=10)

# Add the folder path of note taking
note_label = tk.Label(container, text="Folder Path for Note:", font=("Times New Roman", 17), bg="#b7f1ea")
note_label.grid(row=7, column=0, sticky="w")
note_entry = tk.Entry(container, font=("Times New Roman", 14), width=30)  # Increased width
note_entry.grid(row=7, column=1, padx=10, pady=10)

# Function to handle the button click event
def save_and_register():
    # Perform saving here
    name = name_entry.get()
    email = email_entry.get()
    phone = phone_entry.get()
    city = city_entry.get()
    music = music_entry.get()
    screenshot = screenshot_entry.get()
    note = note_entry.get()
    # Save data to the database
    conn = sqlite3.connect('Desktop_Ai_Assistant.db')
    cursor = conn.cursor()
    cursor.execute("INSERT INTO user (user, data) VALUES (?, ?)", ('name', name))
    cursor.execute("INSERT INTO user (user, data) VALUES (?, ?)", ('email', email))
    cursor.execute("INSERT INTO user (user, data) VALUES (?, ?)", ('phone', phone))
    cursor.execute("INSERT INTO user (user, data) VALUES (?, ?)", ('city', city))
    cursor.execute("INSERT INTO user (user, data) VALUES (?, ?)", ('music', music))
    cursor.execute("INSERT INTO user (user, data) VALUES (?, ?)", ('screenshot', screenshot))
    cursor.execute("INSERT INTO user (user, data) VALUES (?, ?)", ('note', note))
    conn.commit()
    conn.close()
    
    
    
    # if you want to Display data from the 'user' table on the console
    #conn = sqlite3.connect('Desktop_Ai_Assistant.db')
    #cursor = conn.cursor()
    #cursor.execute("SELECT * FROM user")
    #user_data = cursor.fetchall()
    #for row in user_data:
    #   print(f"{row[0]}: {row[1]}")
    #conn.close()
    #print("Data saved successfully. Proceed to register the face.")
    
    
    
    # Creating the folder Called faces
    faces_folder = 'faces'
    if not os.path.exists(faces_folder):
        os.mkdir(faces_folder)
        
    
    
    
    # Get the current working directory (the directory where this script is located)
    script_directory = os.getcwd()

    # Create a batch file in the same directory
    batch_script = """@echo off
    python Train_Test.py
    """

    batch_file_path = os.path.join(script_directory, 'Assistant.bat')

    with open(batch_file_path, 'w') as batch_file:
        batch_file.write(batch_script)

    # Create a shortcut to the batch file on the desktop
    shortcut_name = "Desktop Assistant"

    shell = Dispatch("WScript.Shell")
    desktop_path = shell.SpecialFolders("Desktop")
    shortcut = shell.CreateShortCut(os.path.join(desktop_path, shortcut_name + ".lnk"))
    shortcut.Targetpath = batch_file_path
    shortcut.WorkingDirectory = script_directory  # Set it to the directory where the batch file is located
    #shortcut.IconLocation = 'icon_logo.ico'  # Replace with the actual path to your icon
    shortcut.IconLocation = os.path.join(script_directory, 'icon_logo.ico')
    shortcut.save()


        
    # Execute another Python script
    #script_path = 'D:\Final_Year_Project\Ai\create_data.py'  # Replace with the actual path to your Python script
    script_path = './Collect_Data.py'
    subprocess.Popen(['python', script_path])
    os.system("exit")
    root.destroy()



# Add the "Save and Register Face" button
save_button = tk.Button(container, text="Save and Register Face", bg="#4f8fef", font=("Times New Roman", 16), highlightthickness=2,
                         highlightcolor="black", relief="raised", cursor="hand2", command=save_and_register)
save_button.grid(row=8, column=0, columnspan=2, padx=10, pady=10)



# Configure the container to expand with the window size
container.grid_rowconfigure(0, weight=1)
container.grid_rowconfigure(6, weight=1)
container.grid_columnconfigure(0, weight=1)
container.grid_columnconfigure(2, weight=1)

# Start the Tkinter main loop
root.mainloop()
