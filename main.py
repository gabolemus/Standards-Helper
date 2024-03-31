import tkinter as tk
from tkinter import filedialog, ttk
from typing import Any

from standards.inputs import get_text_comparisons, show_matches
from standards.workbooks import get_new_standards, get_original_standards


class StandardsHelperApp:  # pylint: disable=R0902
    """Main class for the standards helper app."""

    def __init__(self, master):
        self.master = master
        master.title("Standards Helper App")

        # State variables
        self.selected_current_file = False
        self.selected_new_file = False
        self.current_standards = []
        self.selected_standard = None
        self.new_standards = []
        self.keywords = []
        self.potential_new_standards = []

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

        self.filter_entry = tk.Entry(master, width=70)
        self.filter_entry.pack()
        self.filter_entry.bind("<KeyRelease>", lambda _: self.filter_current_standards(
            self.filter_entry.get()))
        self.add_placeholder(self.filter_entry, "Filter by criteria or No.")

        self.keywords_entry = tk.Entry(master, width=70)
        self.keywords_entry.pack()
        self.keywords_entry.bind("<KeyRelease>", lambda _: self.filter_keywords(
            self.keywords_entry.get()))
        self.add_placeholder(self.keywords_entry,
                             "Enter keywords separated by commas")

        self.compare_button = tk.Button(
            master, text="Start Comparison", command=self.start_comparison,
            state="disabled", width=30)
        self.compare_button.pack()

        self.current_standards_tree = None

        self.matching_new_standards_tree = None

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

        self.new_standards = get_new_standards(self.new_file_path)
        print(f"New standards: {len(self.new_standards)}")

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

        if self.current_standards_tree:
            self.current_standards_tree.destroy()

        self.current_standards_tree = ttk.Treeview(
            self.master, columns=("No.", "Criteria", "Level"), show="headings")
        self.current_standards_tree.heading("No.", text="No.")
        self.current_standards_tree.heading("Criteria", text="Criteria")
        self.current_standards_tree.heading("Level", text="Level")
        self.current_standards_tree.column("#0", width=0, stretch=tk.NO)
        self.current_standards_tree.column("No.", width=50, stretch=tk.NO)
        self.current_standards_tree.column(
            "Criteria", width=500, stretch=tk.NO)
        self.current_standards_tree.column("Level", width=100, stretch=tk.NO)
        self.current_standards_tree.pack()

        for standard in self.current_standards:
            if standard:
                self.current_standards_tree.insert("", "end", values=(
                    standard["id"], standard["text"], standard["level"]))

        self.current_standards_tree.bind(
            "<ButtonRelease-1>", self.on_treeview_select)  # Bind the selection event

        print(
            f"Updated current standards. Length: {len(self.current_standards)}")

    def filter_current_standards(self, criteria):
        filtered_standards = []
        if criteria == "":  # If the entry is empty, show all standards
            filtered_standards = self.current_standards
        else:
            for standard in self.current_standards:
                if standard and standard["text"] and standard["id"]:
                    if standard["id"].startswith(criteria) or \
                            criteria in standard["id"].lower() or \
                        standard["text"].lower().startswith(criteria.lower()) or \
                            criteria in standard["text"].lower():
                        filtered_standards.append(standard)

        # Clear the Treeview
        if self.current_standards_tree:
            for item in self.current_standards_tree.get_children():
                self.current_standards_tree.delete(item)

            # Populate the Treeview with filtered standards
            for standard in filtered_standards:
                self.current_standards_tree.insert("", "end", values=(
                    standard["id"], standard["text"], standard["level"]))

    def filter_keywords(self, keywords_input):
        self.keywords = [kw.strip() for kw in keywords_input.split(
            ",") if kw.strip()] if keywords_input else []

    def on_treeview_select(self, _):
        if self.current_standards_tree:
            item = self.current_standards_tree.selection()[0]
            standard = self.current_standards_tree.item(item)["values"]
            self.selected_standard = standard
            print(f"Selected standard: {self.selected_standard}")

    def start_comparison(self):
        # Add logic to compare standards and display results
        print("Starting comparison...")

        if self.selected_standard:
            curr_std = self.selected_standard[0]
            matches: Any = {self.selected_standard[0]: {}}
            original_standard = next(
                (std for std in self.current_standards if std["id"] == self.selected_standard[0]), None)
        else:
            curr_std = "Unknown"
            matches = {}
            original_standard = None

        if original_standard:
            for new_standard in self.new_standards:
                if new_standard["text"] and original_standard["text"]:
                    cosine_sim, edit_dist = get_text_comparisons(
                        original_standard["text"], new_standard["text"])

                    if self.keywords:
                        cosine_weight = 0.3
                        edit_weight = 0.3
                        keyword_weight = 0.4

                        keyword_proportion = len([kw for kw in self.keywords
                                                  if kw in new_standard["text"]]) / len(self.keywords)
                        weighted_similarity = cosine_sim * cosine_weight + edit_dist * \
                            edit_weight + keyword_proportion * keyword_weight

                        matches[curr_std][new_standard["id"]] = {
                            "weighted_similarity": round(100 * weighted_similarity, 2),
                            "cosine": cosine_sim,
                            "edit": edit_dist,
                            "keyword_proportion": keyword_proportion,
                        }
                    else:
                        cosine_weight = 0.5
                        edit_weight = 0.5

                        weighted_similarity = cosine_sim * cosine_weight + \
                            edit_dist * edit_weight

                        matches[curr_std][new_standard["id"]] = {
                            "weighted_similarity": round(100 * weighted_similarity, 2),
                            "cosine": cosine_sim,
                            "edit": edit_dist,
                        }

        show_matches(original_standard, matches, curr_std)

        # Add the matching new standards to self.potential_new_standards
        self.potential_new_standards = []
        # Filter the new standards with the std_id of the matching standards
        # But first, create a new list from the new standards sorted by the
        # weighted similarity
        sorted_new_standards = sorted(self.new_standards, key=lambda x: matches[curr_std].get(
            x["id"], {}).get("weighted_similarity", 0), reverse=True)
        print(f"Sorted new standards: {len(sorted_new_standards)}")
        print(f"First sorted new standard: {sorted_new_standards[0]}")
        for new_standard in sorted_new_standards:
            if new_standard["id"] in matches[curr_std]:
                match = matches[curr_std][new_standard["id"]]
                self.potential_new_standards.append(new_standard)
        print(f"Potential new standards: {len(self.potential_new_standards)}")
        print(
            f"First potential new standard: {self.potential_new_standards[0]}")

        # Display the matching new standards in a Treeview
        if self.matching_new_standards_tree:
            self.matching_new_standards_tree.destroy()

        self.matching_new_standards_tree = ttk.Treeview(
            self.master, columns=("No.", "New Criteria", "Level", "Similarity"), show="headings")
        self.matching_new_standards_tree.heading("No.", text="No.")
        self.matching_new_standards_tree.heading(
            "New Criteria", text="New Criteria")
        self.matching_new_standards_tree.heading("Level", text="Level")
        self.matching_new_standards_tree.heading(
            "Similarity", text="Similarity")
        self.matching_new_standards_tree.column("#0", width=0, stretch=tk.NO)
        self.matching_new_standards_tree.column("No.", width=50, stretch=tk.NO)
        self.matching_new_standards_tree.column(
            "New Criteria", width=500, stretch=tk.NO)
        self.matching_new_standards_tree.column(
            "Level", width=100, stretch=tk.NO)
        self.matching_new_standards_tree.column(
            "Similarity", width=100, stretch=tk.NO)
        self.matching_new_standards_tree.pack()

        # Initially, show only the top 10 matches
        for new_standard in self.potential_new_standards[:10]:
            self.matching_new_standards_tree.insert("", "end", values=(
                new_standard["id"], new_standard["text"], new_standard["level"],
                matches[curr_std].get(new_standard["id"], {}).get("weighted_similarity", 0)))


root = tk.Tk()
root.state("zoomed")
app = StandardsHelperApp(root)
root.mainloop()
