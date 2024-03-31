import tkinter as tk
from tkinter import filedialog, ttk

from standards.workbooks import get_original_standards


class StandardsHelperApp:  # pylint: disable=R0902
    """Main class for the standards helper app."""

    def __init__(self, master):
        self.master = master
        master.title("Standards Helper App")

        # State variables
        self.selected_current_file = False
        self.selected_new_file = False
        self.current_standards = []

        # File paths
        self.current_file_path = ""
        self.new_file_path = ""

        # Worksheets
        self.worksheets = ["RCS", "GRS", "RWS", "RMS", "RAS", "RDS"]
        self.current_worksheet = None
        self.selected_worksheet_index = tk.IntVar(value=-1)

        # UI Elements
        self.current_file_button = tk.Button(
            master, text="Select Comparison File",
            command=self.select_current_file, width=30)
        self.current_file_button.pack()

        self.new_file_button = tk.Button(
            master, text="Select Unified Standards File", command=self.select_new_file, width=30)
        self.new_file_button.pack()

        self.worksheet_frame = tk.Frame(master)
        self.worksheet_frame.pack()

        self.worksheet_radio_buttons = []
        for i, worksheet in enumerate(self.worksheets):
            radio_button = tk.Radiobutton(self.worksheet_frame, text=worksheet,
                                          variable=self.selected_worksheet_index, value=i,
                                          state="disabled", command=self.update_current_worksheet)
            radio_button.pack(side=tk.LEFT)
            self.worksheet_radio_buttons.append(radio_button)

        self.filter_entry = tk.Entry(master, width=30)
        self.filter_entry.pack()
        self.add_placeholder(self.filter_entry, "Filter by name or ID")

        self.keywords_entry = tk.Entry(master, width=30)
        self.keywords_entry.pack()
        self.add_placeholder(self.keywords_entry,
                             "Enter keywords separated by commas")

        self.compare_button = tk.Button(
            master, text="Start Comparison", command=self.start_comparison,
            state="disabled", width=30)
        self.compare_button.pack()

        self.current_standards_tree = None

    def add_placeholder(self, entry, placeholder):
        entry.insert(0, placeholder)
        entry.bind("<FocusIn>", lambda event: self.on_entry_focus_in(
            event, entry, placeholder))
        entry.bind("<FocusOut>", lambda event: self.on_entry_focus_out(
            event, entry, placeholder))

    def on_entry_focus_in(self, _, entry, placeholder):
        if entry.get() == placeholder:
            entry.delete(0, tk.END)
            entry.config(fg="black")  # Change text color to black

    def on_entry_focus_out(self, _, entry, placeholder):
        if entry.get() == "":
            entry.insert(0, placeholder)
            entry.config(fg="gray")  # Change text color to gray

    def select_current_file(self):
        self.current_file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])

        for checkbox in self.worksheet_radio_buttons:
            checkbox.config(state="normal")

        self.selected_worksheet_index.set(0)
        self.current_worksheet = self.worksheets[0]
        self.current_file_button.config(state="disabled")
        self.selected_current_file = True

        if self.selected_new_file:
            self.compare_button.config(state="normal")

        self.display_current_standards()

    def select_new_file(self):
        self.new_file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        self.new_file_button.config(state="disabled")
        self.selected_new_file = True

        if self.selected_current_file:
            self.compare_button.config(state="normal")

    def update_current_worksheet(self):
        index = self.selected_worksheet_index.get()
        self.current_worksheet = self.worksheets[index]

        if self.current_standards:
            self.display_current_standards()

    def display_current_standards(self):
        if not self.current_worksheet:
            worksheet = self.worksheets[0]
        else:
            worksheet = self.current_worksheet

        self.current_standards = get_original_standards(
            self.current_file_path, worksheet)

        # Populate the listbox with the current standards
        if self.current_standards_tree:
            self.current_standards_tree.destroy()

        self.current_standards_tree = ttk.Treeview(
            self.master, columns=("No.", "Criteria", "Level"), show="headings")
        self.current_standards_tree.heading("No.", text="No.")
        self.current_standards_tree.heading("Criteria", text="Criteria")
        self.current_standards_tree.heading("Level", text="Level")
        self.current_standards_tree.pack()

        for standard in self.current_standards:
            if standard:
                self.current_standards_tree.insert("", "end", values=(
                    standard["id"], standard["text"], standard["level"]))

        print(
            f"Updated current standards. Length: {len(self.current_standards)}")

    def start_comparison(self):
        # Add logic to compare standards and display results
        pass


# Start the app maximized
root = tk.Tk()
root.state("zoomed")
app = StandardsHelperApp(root)
root.mainloop()
