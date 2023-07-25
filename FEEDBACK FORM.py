import tkinter as tk
from tkinter import ttk
import pandas as pd
from tkinter import *
def on_configure(event):
    # Update the scrollable region when the frame size changes
    canvas.configure(scrollregion=canvas.bbox("all"))
def on_mousewheel(event):
    # Scroll the canvas when the mouse wheel is moved
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

def submit_feedback():
    #canvas.configure(scrollregion=canvas.bbox("all"))
    name = name_entry.get()
    year_dept = year_dept_entry.get()
    college = college_entry.get()
    mobile = mobile_entry.get()
    whatsapp = whatsapp_entry.get()
    email = email_entry.get()
    facebook = facebook_entry.get()
    information = information_var.get()
    interested = interested_var.get()

    # Create a list of selected courses
    selected_courses = [course_interest_options[i] for i in range(len(course_interest_options)) if
                        course_checkbuttons[i].get()]

    # Create a dictionary from the feedback data
    feedback_data = {
        "Name": name,
        "Year & Department": year_dept,
        "College Name": college,
        "Mobile Number": mobile,
        "WhatsApp Number": whatsapp,
        "Email ID": email,
        "Facebook ID": facebook,
        "Course Interested": ", ".join(selected_courses),
        "information": information,
        "Interested": interested

    }

    # Create a DataFrame from the dictionary
    df = pd.DataFrame([feedback_data])

    # Try to read the existing data from the file if it exists
    try:
        existing_df = pd.read_excel("feedback_data.xlsx")
        df = pd.concat([existing_df, df], ignore_index=True)
    except FileNotFoundError:
        pass

    # Save the DataFrame to the Excel file
    df.to_excel("feedback_data.xlsx", index=False)

    # Clear the form fields after submission
    name_entry.delete(0, tk.END)
    year_dept_entry.delete(0, tk.END)
    college_entry.delete(0, tk.END)
    mobile_entry.delete(0, tk.END)
    whatsapp_entry.delete(0, tk.END)
    email_entry.delete(0, tk.END)
    facebook_entry.delete(0, tk.END)
    information_var.set("")
    interested_var.set("")

    for i in range(len(course_interest_options)):
        course_checkbuttons[i].set(False)


root = tk.Tk()
root.title("Feedback Form")
custom_font = ("Arial", 18, "bold")
name_label = ttk.Label(root, text="STUDENT FEEDBACK FORM", font=custom_font)
name_label.pack()

# Create a vertical scrollbar
scrollbar = tk.Scrollbar(root)
scrollbar.pack(side="right", fill="y")

#Create a frame to hold the scrollable content
canvas = tk.Canvas(root, yscrollcommand=scrollbar.set)
canvas.pack()
# Link the scrollbar to the canvas
scrollbar.config(command=canvas.yview)

frame = tk.Frame(root)
canvas.create_window((0, 0), window=frame, anchor='nw')

i = 0
j = 20
k = 25
label_padding = 10
# Name Lable
i += 1
name_label = ttk.Label(frame, text="Name:")
name_label.grid(row=i, column=j, sticky=tk.W, padx=label_padding, pady=label_padding)

name_entry = ttk.Entry(frame)
name_entry.grid(row=i, column=k, sticky=tk.W, padx=label_padding, pady=label_padding)

i += 1
year_dept_label = ttk.Label(frame, text="Year & Department:")
year_dept_label.grid(row=i, column=j, sticky=tk.W, padx=label_padding, pady=label_padding)
year_dept_entry = ttk.Entry(frame)
year_dept_entry.grid(row=i, column=k, sticky=tk.W, padx=label_padding, pady=label_padding)

i += 1

college_label = ttk.Label(frame, text="College Name:")
college_label.grid(row=i, column=j, sticky=tk.W, padx=label_padding, pady=label_padding)
college_entry = ttk.Entry(frame)
college_entry.grid(row=i, column=k, sticky=tk.W, padx=label_padding, pady=label_padding)

i += 1

mobile_label = ttk.Label(frame, text="Mobile Number:")
mobile_label.grid(row=i, column=j, sticky=tk.W, padx=label_padding, pady=label_padding)
mobile_entry = ttk.Entry(frame)
mobile_entry.grid(row=i, column=k, sticky=tk.W, padx=label_padding, pady=label_padding)

i += 1

whatsapp_label = ttk.Label(frame, text="WhatsApp Number:")
whatsapp_label.grid(row=i, column=j, sticky=tk.W, padx=label_padding, pady=label_padding)
whatsapp_entry = ttk.Entry(frame)
whatsapp_entry.grid(row=i, column=k, sticky=tk.W, padx=label_padding, pady=label_padding)

i += 1

email_label = ttk.Label(frame, text="Email ID:")
email_label.grid(row=i, column=j, sticky=tk.W, padx=label_padding, pady=label_padding)
email_entry = ttk.Entry(frame)
email_entry.grid(row=i, column=k, sticky=tk.W, padx=label_padding, pady=label_padding)

i += 1

facebook_label = ttk.Label(frame, text="Facebook ID:")
facebook_label.grid(row=i, column=j, sticky=tk.W, padx=label_padding, pady=label_padding)
facebook_entry = ttk.Entry(frame)
facebook_entry.grid(row=i, column=k, sticky=tk.W, padx=label_padding, pady=label_padding)

i += 1
information_label = ttk.Label(frame, text="Do you Want to get more information about this field")
information_label.grid(row=i, column=j, sticky=tk.W, padx=label_padding, pady=label_padding)

i += 1
information_var = tk.StringVar(value="Yes")
information_yes_radio = ttk.Radiobutton(frame, text="Yes", variable=information_var, value="Yes")
information_yes_radio.grid(row=i, column=j, sticky=tk.W, padx=label_padding, pady=label_padding)

information_no_radio = ttk.Radiobutton(frame, text="No", variable=information_var, value="No")
information_no_radio.grid(row=i, column=k, sticky=tk.W, padx=label_padding, pady=label_padding)

i += 1
interested_label = ttk.Label(frame, text="Are you interested to join the course now?")
interested_label.grid(row=i, column=j, sticky=tk.W, padx=label_padding, pady=label_padding)

i += 1
interested_var = tk.StringVar(value="Yes")
interested_yes_radio = ttk.Radiobutton(frame, text="Yes", variable=interested_var, value="Yes")
interested_yes_radio.grid(row=i, column=j, sticky=tk.W, padx=label_padding, pady=label_padding)

interested_no_radio = ttk.Radiobutton(frame, text="No", variable=interested_var, value="No")
interested_no_radio.grid(row=i, column=k, sticky=tk.W, padx=label_padding, pady=label_padding)

i += 1

course_interest_label = ttk.Label(frame, text="Which course are you interested to join?")
course_interest_label.grid(row=i, column=j, sticky=tk.W, padx=label_padding, pady=label_padding)
course_interest_options = [
    "Python", "NodeJS", "React JS", "Back end development",
    "Front end development", "Bootstrap programming",
    "JavaScript HTML-5 & CSS-3",
    "Full stack development using Python/MEAN/MERN",
    "Web designing", "Django"
]
course_checkbuttons = []

i += 1

for course in course_interest_options:
    var = tk.BooleanVar()
    course_checkbutton = ttk.Checkbutton(frame, text=course, variable=var)
    course_checkbutton.grid(row=i, column=j, sticky=tk.W, padx=label_padding, pady=label_padding)
    i = i + 1
    course_checkbuttons.append(var)

i += 1
submit_button = ttk.Button(frame, text="Submit", width=20, command=submit_feedback)
submit_button.grid(row=i, column=j, sticky=(tk.W, tk.E, tk.N, tk.S))

# Add widgets to the frame (scrollable content)
for i in range(5):
    tk.Label(frame).grid(row=i, column=0)

# Configure the frame to adjust its size when the canvas size changes
frame.bind("<Configure>", on_configure)
canvas.bind_all("<MouseWheel>", on_mousewheel)
# Make the grid cell grow with the window size
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

root.mainloop()