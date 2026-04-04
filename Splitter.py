# ============================================================
# Project Name : Splitter
# Developed By : Mounir Boudhan
# Description  : This desktop application was developed to
#                automate the process of converting a Word
#                document (.docx) into multiple separate PDF
#                files, where each page is extracted and saved
#                individually.
#
# Purpose      : The main reason behind developing this app is
#                to save time, reduce manual work, and simplify
#                document processing for files that contain
#                multiple pages requiring separate PDF outputs.
# ============================================================

from tkinter import *

win = Tk()

# Set window size
window_width = 780
window_height = 500

# Get screen size
screen_width = win.winfo_screenwidth()
screen_height = win.winfo_screenheight()

# Calculate center position
x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)

# Apply window size and position
win.geometry(f"{window_width}x{window_height}+{x}+{y}")

# Disable window resizing
win.resizable(False, False)

# Set application window title
win.title("Splitter")

# Set application window icon
######### win.iconbitmap("icon.ico")


win.mainloop()