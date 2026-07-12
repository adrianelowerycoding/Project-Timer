import os
import openpyxl
import tkinter as tk

from tkinter import ttk
from tkinter import filedialog, simpledialog, messagebox
from pathlib import Path

#### Creating the GUI ####

# -- Dark mode palette: tweak these hex values to taste --
BG = "#1e1e1e"          # main window / frame background
BG_ALT = "#2b2b2b"      # slightly lighter background (boxes, buttons)
FG = "#e0e0e0"          # primary text color
FG_MUTED = "#a0a0a0"    # secondary/less prominent text
ACCENT = "#4CAF50"      # accent color (e.g. idle circle button)
BORDER = "#3c3c3c"      # borders/separators
DELETE_RED = "#c0392b"  # color for the "x" delete button


# The timer you create per project
class TimerBox(tk.Frame):
    def __init__(self, parent, project_name, on_delete, **kwargs):
        # A slightly bigger frame so it reads as a "medium-sized square" button
        super().__init__(
            parent, bd=1, relief="solid", padx=10, pady=10,
            bg=BG_ALT, highlightbackground=BORDER,
            width=140, height=140, **kwargs
        )
        # Keep the frame from shrinking to fit its children so it stays square-ish
        self.pack_propagate(False)
        self.grid_propagate(False)

        self.project_name = project_name
        self.on_delete = on_delete  # callback the MainWindow gives us: on_delete(self)

        # -- "x" delete button, placed in the top-right corner --
        # Hidden by default; MainWindow toggles visibility via show_delete_x()/hide_delete_x()
        self.delete_btn = tk.Button(
            self, text="x", font=("Segoe UI", 8, "bold"),
            bg=DELETE_RED, fg="white", activebackground="#e74c3c", activeforeground="white",
            relief="flat", highlightthickness=0, bd=0,
            width=2, height=1, command=self.confirm_delete
        )
        # place() lets it float in the corner without disturbing the pack layout below

        self.title_label = tk.Label(
            self, text=project_name, font=("Segoe UI", 11, "bold"),
            bg=BG_ALT, fg=FG, wraplength=110, justify="center"
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

    def show_delete_x(self):
        # Anchor to the top-right corner of the frame
        self.delete_btn.place(relx=1.0, x=-2, y=2, anchor="ne")

    def hide_delete_x(self):
        self.delete_btn.place_forget()

    def confirm_delete(self):
        answer = messagebox.askyesno(
            title="Delete Timer",
            message=f'Do you really want to delete "{self.project_name}"\'s timer?'
        )
        if answer:
            self.on_delete(self)


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
        self.current_database_name = tk.StringVar(value="Current Database: None")  # Initial text

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
        self.create_db_btn = create_db_btn  # kept so we can disable it during delete mode

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
        self.options_btn = options_btn  # kept so we can disable it during delete mode

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
        self.create_timer_btn = create_timer_btn  # kept so we can disable it during delete mode

        # Keep a reference to this button so we can re-label it Done/Delete Timer
        self.delete_timer_btn = tk.Button(
            frame, text="Delete Timer", command=self.on_delete_timer,
            bg=BG_ALT, fg=FG, activebackground=BORDER, activeforeground=FG,
            relief="flat", highlightthickness=0
        )
        self.delete_timer_btn.pack(side="left")

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

        if user_input_project_numbers is None:
            return  # Cancelled - stop here, don't open save dialog
        if user_input_project_numbers.strip() == "":
            return  # Input was blank - re-prompt instead of just quitting

        # Strips the commas and spaces from project numbers
        user_input_project_numbers = user_input_project_numbers.replace(",", "_").replace(" ", "")
        # TODO

        # Captures username and sets path for suggested_folder and creates it
        username = os.environ.get("USERNAME", "user")  # Captures Window's username
        suggested_folder = Path(os.environ.get("USERPROFILE", r"C:\Users\Default")) / "Documents" / "Project Tracker" / f"db {user_input_project_numbers}"
        suggested_folder.mkdir(parents=True, exist_ok=True)  # creates folder

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
        self.set_current_database(self.current_database)  # Passes database name into method

        return Path(filepath)

    # Updates the "Curent Database: <database>" title from 'self.current_database' in
    # 'on_create_database()'. Showcases name in Title
    def set_current_database(self, name):
        self.current_database_name.set(f"Current Database: {name}")

    def on_update_database(self):
        pass  # TODO

    def on_view_days_data(self):
        pass  # TODO

    # -- Creates a new TimerBox --
    def on_create_project_timer(self):
        project_name = simpledialog.askstring(
            title="New Project Timer",
            prompt="What project is this timer for?"
        )

        if project_name is None:
            return  # Cancelled
        if project_name.strip() == "":
            return  # Blank input, ignore

        project_name = project_name.strip()

        timer = TimerBox(self.timers_container, project_name, on_delete=self.remove_timer)
        self.timers.append(timer)
        self._relayout_timers()

    # -- Toggles delete mode: shows/hides the "x" on every timer box, and disables
    #    every other button so "Done Deleting" is the only way out --
    def on_delete_timer(self):
        self.delete_mode = not self.delete_mode

        for timer in self.timers:
            if self.delete_mode:
                timer.show_delete_x()
            else:
                timer.hide_delete_x()

        # Stored as 'self' attributes so now reachable from 'on_delete_timer()' 
        other_state = tk.DISABLED if self.delete_mode else tk.NORMAL
        self.create_db_btn.config(state=other_state)
        self.options_btn.config(state=other_state)
        self.create_timer_btn.config(state=other_state)

        self.delete_timer_btn.config(text="Done Deleting" if self.delete_mode else "Delete Timer")

    # -- Called by a TimerBox once the user confirms deletion --
    def remove_timer(self, timer):
        if timer in self.timers:
            self.timers.remove(timer)
        timer.destroy()
        self._relayout_timers()

    # -- Arranges all current TimerBoxes into a grid inside timers_container --
    def _relayout_timers(self):
        for widget in self.timers_container.winfo_children():
            widget.grid_forget()

        for index, timer in enumerate(self.timers):
            row = index // self.COLUMNS
            col = index % self.COLUMNS
            timer.grid(row=row, column=col, padx=8, pady=8)


if __name__ == "__main__":
    mainWindow = MainWindow()
    # Begins an event loop, listening for OS-level events and triggering callback functions
    mainWindow.mainloop()