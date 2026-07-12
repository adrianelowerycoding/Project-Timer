import os
import openpyxl
import tkinter as tk 
from tkinter import ttk
from tkinter import filedialog, simpledialog
from pathlib import Path


#### Creating the GUI #### 

# -- Dark mode palette: tweak these hex values to taste --
BG = "#1e1e1e"          # main window / frame background
BG_ALT = "#2b2b2b"      # slightly lighter background (boxes, buttons)
FG = "#e0e0e0"          # primary text color
FG_MUTED = "#a0a0a0"    # secondary/less prominent text
ACCENT = "#4CAF50"      # accent color (e.g. idle circle button)
BORDER = "#3c3c3c"      # borders/separators


# The timer you create per project
class TimerBox(tk.Frame):
    def __init__(self, parent, project_number, **kwargs):
        super().__init__(
            parent, bd=1, relief="solid", padx=10, pady=10,
            bg=BG_ALT, highlightbackground=BORDER, **kwargs
        )

        self.project_number = project_number

        self.title_label = tk.Label(
            self, text=f"Project {project_number}", font=("Segoe UI", 11, "bold"),
            bg=BG_ALT, fg=FG
        )
        self.title_label.pack(pady=(0, 6))

        self.canvas = tk.Canvas(self, width=70, height=70, highlightthickness=0, bg=BG_ALT)
        self.canvas.pack()

        self.circle = self.canvas.create_oval(4, 4, 66, 66, fill=ACCENT, outline="")
        self.circle_text = self.canvas.create_text(35, 35, text="Start", fill="white", font=("Segoe UI", 10, "bold"))
        self.canvas.bind("<Button-1>", self.on_toggle_clicked)

        self.time_label = tk.Label(self, text="00:00:00", font=("Consolas", 10), bg=BG_ALT, fg=FG_MUTED)
        self.time_label.pack(pady=(6, 0))

    def on_toggle_clicked(self, event=None):
        pass  # TODO: your start/stop logic here

# The main window
class MainWindow(tk.Tk):

    # __init__ gets called as soon as mainWindow() object is called
    def __init__(self):
        super().__init__()
        self.title("Project Time Tracker")
        self.geometry("480x600")
        self.minsize(400, 500)
        self.configure(bg=BG)
        self.current_database = None
        self.current_database_name = tk.StringVar(value="Current Database: None") # Initial text


        # ttk widgets (just the Separators here) need a Style, since bg/fg
        # kwargs don't work on them directly like they do on plain tk widgets.
        style = ttk.Style(self)
        style.theme_use("clam")  # 'clam' respects custom colors; 'default'/'vista' mostly ignore them
        style.configure("TSeparator", background=BG)

        self.COLUMNS = 3

        # Keeps track of every TimerBox currently on screen
        self.timers = []
        # Whether the "x" delete buttons are currently showing on the timers
        self.delete_mode = False

        # Runs every single function inside mainWindow()
        # Creates the buttons and registers their references/callbacks
        self._build_section1()
        self._build_section2()
        self._build_title_bar()
        self._build_section3()

    # -- SECTION 1: Database button + Options dropdown --
    # Creates the button(s) and hands Tkinter a reference to run logic once the button is pressed
    def _build_section1(self):
        frame = tk.Frame(self, pady=10, bg=BG)
        frame.pack(fill="x", padx=10)

        create_db_btn = tk.Button(
            frame, text="Create Database", command=self.on_create_database,
            bg=BG_ALT, fg=FG, activebackground=BORDER, activeforeground=FG,
            relief="flat", highlightthickness=0
        )
        create_db_btn.pack(side="left", padx=(0, 10))

        options_btn = tk.Menubutton(
            frame, text="Options", relief="flat", indicatoron=True,
            bg=BG_ALT, fg=FG, activebackground=BORDER, activeforeground=FG,
            highlightthickness=0
        )
        options_menu = tk.Menu(
            options_btn, tearoff=False, bg=BG_ALT, fg=FG,
            activebackground=BORDER, activeforeground=FG
        )
        options_menu.add_command(label="Update Database", command=self.on_update_database)
        options_menu.add_command(label="View Day's Data", command=self.on_view_days_data)
        options_btn.config(menu=options_menu)
        options_btn.pack(side="left")

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=5, pady=(0, 5))

    # -- SECTION 2: Create/Delete timer buttons --
    def _build_section2(self):
        frame = tk.Frame(self, pady=5, bg=BG)
        frame.pack(fill="x", padx=10)

        create_timer_btn = tk.Button(
            frame, text="Create Project Timer", command=self.on_create_project_timer,
            bg=BG_ALT, fg=FG, activebackground=BORDER, activeforeground=FG,
            relief="flat", highlightthickness=0
        )
        create_timer_btn.pack(side="left", padx=(0, 10))

        delete_timer_btn = tk.Button(
            frame, text="Delete Timer", command=self.on_delete_timer,
            bg=BG_ALT, fg=FG, activebackground=BORDER, activeforeground=FG,
            relief="flat", highlightthickness=0
        )
        delete_timer_btn.pack(side="left")

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=5, pady=(5, 5))

    # -- Dynamic title bar: shows current database name --
    def _build_title_bar(self):
        label = tk.Label(
            self, textvariable=self.current_database_name,
            font=("Segoe UI", 9, "italic"), bg=BG, fg=FG_MUTED
        )
        label.pack(fill="x", padx=10, pady=(0, 5))
 

    # -- SECTION 3: Scrollable box of timer widgets --
    def _build_section3(self):
        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.canvas = tk.Canvas(outer, highlightthickness=0, bg=BG)
        scrollbar = tk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.timers_container = tk.Frame(self.canvas, bg=BG)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.timers_container, anchor="nw")

        self.timers_container.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))

    # -- Callback stubs --
    # All of the logic behind the buttons live in these callback functions

    # User creates database
    def on_create_database(self):

        # User inputs all project #s going into the database
        user_input_project_numbers = simpledialog.askstring(
            title="Input Project Numbers",
            prompt="Enter all project numbers included for this database (sep. by commas):")
        
        # Strips the commas and spaces from project numbers
        user_input_project_numbers = user_input_project_numbers.replace(",", "_").replace(" ", "")
        if user_input_project_numbers is None: 
            return # Cancelled - stop here, don't open save dialog
        if user_input_project_numbers.strip() == "":
            return # Input was blank - re-prompt instead of just quitting
            # TODO
       
        # Captures username and sets path for suggested_folder and creates it
        username = os.environ.get("USERNAME", "user") # Captures Window's username
        suggested_folder = Path(os.environ.get("USERPROFILE", r"C:\Users\Default")) / "Documents" / "Project Tracker" / f"db {user_input_project_numbers}"
        suggested_folder.mkdir(parents=True, exist_ok=True) # creates folder

        # Opens File Explorer into created folder and offers save options
        filepath = filedialog.asksaveasfilename(
            title="Choose location for your Time Tracking database",
            initialdir=suggested_folder,
            initialfile=f"{username}_Project_Tracker_db_{user_input_project_numbers}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")]
        )

        if not filepath: 
            return None
        
        # Setting Current Database Name and sending it to set_current_database() so it can 
        # display current database name in the Title
        self.current_database = f"{username}_Project_Tracker_db_{user_input_project_numbers}.xlsx"
        self.set_current_database(self.current_database) # Passes database name into method
    
        return Path(filepath)
    
        
    # Updates the "Curent Database: <database>" title from 'self.current_database' in
    # 'on_create_database()'. Showcases name in Title
    def set_current_database(self, name):
        self.current_database_name.set(f"Current Database: {name}")

    def on_update_database(self):
        pass  # TODO

    def on_view_days_data(self):
        pass  # TODO

    def on_create_project_timer(self):
        pass  # TODO

    def on_delete_timer(self):
        pass  # TODO


if __name__ == "__main__":
    mainWindow = MainWindow()
    # Begins an event loop, listening for OS-level events and triggering callback functions
    mainWindow.mainloop()



#### Getting Folder and File Name ######

# Start out asking the user to first save their database to the folder of their choice and 
# file name of their chocie (auto-suggest a file name)

# Auto-detect Windows username to help create the path

# Set the Excel database path

# If Excel database doesn't exist yet, create one for user. 
# Somehow set this path as the new Excel_Database constant


# Folder/File name = "C:\Users\Excel\{username}_time_log.xsxl"

# create Project Timer folder for user in documents and open up in there to suggest they save in
# there





