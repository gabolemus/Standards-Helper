import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from standards.inputs import get_text_comparisons, show_matches
from standards.workbooks import (get_new_standards, get_original_standards,
                                 update_standards)


class StandardsHelperApp:  # pylint: disable=R0902
    """Main class for the standards helper app."""

    def __init__(self, master):  # pylint: disable=R0915
        self.master = master
        master.title("Standards Helper")

        # State variables
        self.selected_current_file = False
        self.selected_new_file = False
        self.current_standards = []
        self.filtered_standards = []
        self.selected_standard = None
        self.filtered_new_standards = []
        self.new_standards = []
        self.keywords = []
        self.matches = {}
        self.potential_new_standards = []
        self.selected_new_standard = None

        # File paths
        self.current_file_path = ""
        self.new_file_path = ""

        # Worksheets
        self.worksheets = ["RCS", "GRS", "RWS", "RMS", "RAS", "RDS"]
        self.current_worksheet = None
        self.selected_worksheet_index = tk.IntVar(value=-1)

        # UI Elements
        self.central_frame = tk.Frame(master)
        self.central_frame.pack()
        self.left_frame = tk.Frame(self.central_frame)
        self.left_frame.pack(side=tk.LEFT, padx=10)

        self.file_frame = tk.Frame(self.left_frame)
        self.file_frame.pack()

        self.current_file_button = tk.Button(
            self.file_frame, text="Select Comparison File",
            command=self.select_current_file, width=34, font=("Calibri", 11))
        self.current_file_button.pack(side=tk.LEFT)
        self.current_file_button.config(bg="#b3ffb3")
        self.current_file_button.bind(
            "<Enter>", lambda event,
            button=self.current_file_button: self.change_cursor(event, button))
        self.current_file_button.bind(
            "<Leave>", lambda _: self.master.config(cursor=""))

        self.new_file_button = tk.Button(
            self.file_frame, text="Select Unified Standards File",
            command=self.select_new_file, width=34, font=("Calibri", 11))
        self.new_file_button.pack(side=tk.LEFT, padx=5)
        self.new_file_button.config(bg="#b3e6ff")
        self.new_file_button.bind(
            "<Enter>", lambda event,
            button=self.new_file_button: self.change_cursor(event, button))
        self.new_file_button.bind(
            "<Leave>", lambda _: self.master.config(cursor=""))

        self.worksheet_frame = tk.Frame(self.left_frame)
        self.worksheet_frame.pack(pady=10)

        self.worksheets_label = tk.Label(
            self.worksheet_frame, text="Worksheets:", font=("Calibri", 11))
        self.worksheets_label.pack(pady=5)

        self.worksheet_radio_buttons = []
        for i, worksheet in enumerate(self.worksheets):
            radio_button = tk.Radiobutton(self.worksheet_frame, text=worksheet,
                                          variable=self.selected_worksheet_index, value=i,
                                          state="disabled", command=self.update_current_worksheet,
                                          font=("Calibri", 11))
            radio_button.pack(side=tk.LEFT)
            self.worksheet_radio_buttons.append(radio_button)

        self.filter_entry = tk.Entry(
            self.left_frame, width=70, fg="gray", font=("Calibri", 11))
        self.filter_entry.pack(pady=5)
        self.filter_entry.bind("<KeyRelease>", lambda _: self.filter_current_standards(
            self.filter_entry.get()))
        self.add_placeholder(
            self.filter_entry, "Enter the name of the criteria or its No.")

        self.keywords_entry = tk.Entry(
            self.left_frame, width=70, fg="gray", font=("Calibri", 11))
        self.keywords_entry.pack(pady=5)
        self.keywords_entry.bind("<KeyRelease>", lambda _: self.filter_keywords(
            self.keywords_entry.get()))
        self.add_placeholder(self.keywords_entry,
                             "Enter keywords separated by commas")

        self.compare_button = tk.Button(
            self.left_frame, text="Start Comparison", command=self.start_comparison,
            state="disabled", width=70, font=("Calibri", 11))
        self.compare_button.pack(pady=5)
        self.compare_button.config(bg="#b3ffb3")
        self.compare_button.bind(
            "<Enter>", lambda event,
            button=self.compare_button: self.change_cursor(event, button))
        self.compare_button.bind(
            "<Leave>", lambda _: self.master.config(cursor=""))

        self.right_frame = tk.Frame(self.central_frame)
        self.right_frame.pack(side=tk.RIGHT, padx=10)

        self.program_requirements = ttk.Treeview(
            self.right_frame, columns=("Status"), show="headings")
        self.program_requirements.heading("#0", text="Program Requirements")
        self.program_requirements.heading("Status", text="Status")
        self.program_requirements.column("Status", width=325, stretch=tk.NO)
        self.program_requirements.pack(fill="both", expand=True)
        self.program_requirements.bind("<Button-1>", lambda _: "break")

        # Define requirements
        self.obligatory_requirements = {
            'Loaded Comparisons file': False,
            'Loaded Unified Standards file': False,
            'Chosen a standard to compare': False
        }

        self.optional_requirements = {
            'Entered criteria name/No.': False,
            'Entered keyword(s)': False
        }

        self.populate_tree()

        self.helper_inputs_frame = tk.Frame(master)
        self.helper_inputs_frame.pack(pady=10)

        self.currently_selected_frame = tk.Frame(self.helper_inputs_frame)
        self.currently_selected_frame.pack(side=tk.LEFT, padx=50)
        self.currently_selected_label = None
        self.currently_selected_label_text = tk.StringVar()

        self.show_only_non_matched = tk.IntVar(value=0)
        self.show_only_non_completed_standards = None

        self.current_stds_frame = None
        self.current_standards_tree = None
        self.current_stds_scrollbar = None
        self.process_next_std_button = None

        self.filter_new_stds_entry = None
        self.new_stds_frame = None
        self.matching_new_standards_tree = None
        self.new_stds_scrollbar = None

        self.more_matches_frame = None
        self.show_more_matches_button = None
        self.show_all_matches_button = None

        self.write_selected_new_standard_button = None
        self.reset_button = None

    def populate_tree(self):
        # If the tree has items in it, clear them
        for item in self.program_requirements.get_children():
            self.program_requirements.delete(item)

        # Insert obligatory requirements into the requirements tree
        self.program_requirements.insert(
            "", "end", values=("Obligatory Requirements", ""))
        for requirement, status in self.obligatory_requirements.items():
            self.program_requirements.insert(
                "", "end", values=(requirement, status), tags=(self.get_status_color(status),))

        # Insert optional requirements into the requirements tree
        self.program_requirements.insert(
            "", "end", values=("", ""))
        self.program_requirements.insert(
            "", "end", values=("Optional Requirements", ""))
        for requirement, status in self.optional_requirements.items():
            self.program_requirements.insert(
                "", "end", values=(requirement, status),
                tags=(self.get_status_color(status, True),))

        # Color the requirements tree with the appropriate colors
        self.program_requirements.tag_configure('green', background='#b3ffb3')
        self.program_requirements.tag_configure('yellow', background='#ffffb3')
        self.program_requirements.tag_configure('red', background='#ff9999')

    def get_status_color(self, status, is_optional=False):
        if is_optional:
            if status:
                return 'green'
            return 'yellow'

        if status:
            return 'green'
        return 'red'

    def add_placeholder(self, entry, placeholder):
        entry.insert(0, placeholder)
        entry.bind("<FocusIn>", lambda event: self.on_entry_focus_in(
            event, entry, placeholder))
        entry.bind("<FocusOut>", lambda event: self.on_entry_focus_out(
            event, entry, placeholder))

    def on_entry_focus_in(self, _, entry, placeholder):
        if entry.get() == placeholder:
            entry.delete(0, tk.END)
            entry.config(fg="black")

    def on_entry_focus_out(self, _, entry, placeholder):
        if entry.get() == "":
            entry.insert(0, placeholder)
            entry.config(fg="gray")

    def change_cursor(self, _, button):
        if button['state'] == tk.NORMAL:
            self.master.config(cursor="hand2")
        else:
            self.master.config(cursor="")

    def select_current_file(self):
        self.current_file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])

        for checkbox in self.worksheet_radio_buttons:
            checkbox.config(state="normal")

        self.selected_worksheet_index.set(0)
        self.current_worksheet = self.worksheets[0]
        self.current_file_button.config(state="disabled")
        self.selected_current_file = True

        if self.selected_new_file and self.selected_standard:
            self.compare_button.config(state="normal")

        # Update the currently selected standard label
        if self.currently_selected_label:
            self.currently_selected_label.destroy()

        self.currently_selected_label = tk.Label(
            self.currently_selected_frame, textvariable=self.currently_selected_label_text,
            font=("Calibri", 11))
        self.currently_selected_label.pack()

        if self.selected_standard:
            self.currently_selected_label_text.set(
                f"Currently selected standard: {self.selected_standard[0]}")
        else:
            self.currently_selected_label_text.set(
                "Currently selected standard: None")

        # Add the checkbox to show only non-matched standards
        if self.show_only_non_completed_standards:
            self.show_only_non_completed_standards.destroy()

        self.show_only_non_completed_standards = tk.Checkbutton(
            self.helper_inputs_frame, text="Show only non-complete standards",
            variable=self.show_only_non_matched, font=("Calibri", 11),
            command=self.show_only_non_complete_standards)
        self.show_only_non_completed_standards.pack(side=tk.RIGHT, padx=50)

        self.display_current_standards()

        # Update the requirements tree
        self.obligatory_requirements['Loaded Comparisons file'] = True
        self.populate_tree()

    def show_only_non_complete_standards(self):
        if self.show_only_non_matched.get():
            self.filtered_standards = [
                std for std in self.current_standards if not std["completed"]]
        else:
            self.filtered_standards = self.current_standards

        if self.current_standards_tree:
            for item in self.current_standards_tree.get_children():
                self.current_standards_tree.delete(item)

            for standard in self.filtered_standards:
                self.current_standards_tree.insert("", "end", values=(
                    standard["id"], standard["text"], standard["level"],
                    "✓" if standard["completed"] else "✗"))

    def select_new_file(self):
        self.new_file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        self.new_file_button.config(state="disabled")
        self.selected_new_file = True

        if self.selected_current_file and self.selected_standard:
            self.compare_button.config(state="normal")

        self.new_standards = get_new_standards(self.new_file_path)
        # print(f"New standards: {len(self.new_standards)}")

        # Update the requirements tree
        self.obligatory_requirements['Loaded Unified Standards file'] = True
        self.populate_tree()

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

        if self.current_stds_frame:
            self.current_stds_frame.destroy()

        if self.current_stds_scrollbar:
            self.current_stds_scrollbar.destroy()

        self.current_stds_frame = tk.Frame(self.master)
        self.current_stds_frame.pack(pady=10)

        self.current_standards_tree = ttk.Treeview(
            self.current_stds_frame, columns=("No.", "Criteria", "Level", "Completed"), show="headings",
            selectmode="browse")

        style = ttk.Style()
        style.configure("Treeview", font=("Calibri", 11))
        style.configure("Treeview.Heading", font=("Calibri", 11, "bold"))

        self.current_standards_tree.heading("No.", text="No.")
        self.current_standards_tree.heading("Criteria", text="Criteria")
        self.current_standards_tree.heading("Level", text="Level")
        self.current_standards_tree.heading("Completed", text="Completed")
        self.current_standards_tree.column("#0", width=0, stretch=tk.NO)
        self.current_standards_tree.column("No.", width=50, stretch=tk.NO)
        self.current_standards_tree.column(
            "Criteria", width=750, stretch=tk.NO)
        self.current_standards_tree.column("Level", width=100, stretch=tk.NO)
        self.current_standards_tree.column(
            "Completed", width=100, stretch=tk.NO)
        self.current_standards_tree.pack(
            in_=self.current_stds_frame, side="left", fill="both", expand=True)

        self.current_stds_scrollbar = ttk.Scrollbar(
            self.current_stds_frame, orient="vertical", command=self.current_standards_tree.yview)
        self.current_standards_tree.configure(
            yscrollcommand=self.current_stds_scrollbar.set)
        self.current_stds_scrollbar.pack(side="right", fill="y")

        for standard in self.current_standards:
            if standard:
                self.current_standards_tree.insert("", "end", values=(
                    standard["id"], standard["text"], standard["level"],
                    "✓" if standard["completed"] else "✗"))

        self.current_standards_tree.bind(
            "<ButtonRelease-1>", self.on_treeview_select)

        # print(f"Updated current standards. Length: {len(self.current_standards)}")

    def process_next_standard(self):
        if self.filtered_standards and self.current_standards_tree:
            self.filtered_standards.pop(0)

            if self.current_standards_tree:
                self.current_standards_tree.delete(
                    self.current_standards_tree.get_children()[0])

            if self.filtered_standards:
                self.current_standards_tree.selection_set(
                    self.current_standards_tree.get_children()[0])
                self.selected_standard = self.current_standards_tree.item(
                    self.current_standards_tree.get_children()[0])["values"]
                self.start_comparison()
            else:
                self.show_popup(
                    "End of List", "You have reached the end of the list.")

    def filter_current_standards(self, criteria):
        self.filtered_standards = []
        if criteria == "":  # If the entry is empty, show all standards
            self.filtered_standards = self.current_standards
        else:
            for standard in self.current_standards:
                if standard and standard["text"] and standard["id"]:
                    if standard["id"].startswith(criteria) or \
                            criteria in standard["id"].lower() or \
                        standard["text"].lower().startswith(criteria.lower()) or \
                            criteria in standard["text"].lower():
                        self.filtered_standards.append(standard)

        # Check if only the non-completed standards should be shown
        if self.show_only_non_matched.get():
            self.filtered_standards = [
                std for std in self.filtered_standards if not std["completed"]]

        # Enable the process_next_std_button if there are more than 1 standards
        if len(self.filtered_standards) > 1 and self.process_next_std_button:
            if self.process_next_std_button:
                self.process_next_std_button.config(state="normal")

        # Clear the Treeview
        if self.current_standards_tree:
            for item in self.current_standards_tree.get_children():
                self.current_standards_tree.delete(item)

            # Populate the Treeview with filtered standards
            for standard in self.filtered_standards:
                self.current_standards_tree.insert("", "end", values=(
                    standard["id"], standard["text"], standard["level"],
                    "✓" if standard["completed"] else "✗"))

        # Update the requirements tree
        self.optional_requirements['Entered criteria name/No.'] = bool(
            criteria)
        self.populate_tree()

    def filter_new_standards(self, criteria):
        # This function filters the new standards based on the criteria entered
        # in the filter_new_stds_entry. The criteria may match the ID or the text
        # of the new standard.
        self.filtered_new_standards = []
        if criteria == "":  # If the entry is empty, show all standards
            self.filtered_new_standards = self.potential_new_standards
        else:
            for standard in self.potential_new_standards:
                if standard and standard["text"] and standard["id"]:
                    if standard["id"].startswith(criteria) or \
                            criteria in standard["id"].lower() or \
                        standard["text"].lower().startswith(criteria.lower()) or \
                            criteria in standard["text"].lower():
                        self.filtered_new_standards.append(standard)

        # Clear the Treeview
        if self.matching_new_standards_tree and self.selected_standard:
            for item in self.matching_new_standards_tree.get_children():
                self.matching_new_standards_tree.delete(item)

            # Populate the Treeview with filtered standards
            for standard in self.filtered_new_standards:
                self.matching_new_standards_tree.insert("", "end", values=(
                    standard["id"], standard["text"], standard["level"],
                    self.matches[self.selected_standard[0]].get(
                        standard["id"], {}).get("weighted_similarity", 0)))

    def filter_keywords(self, keywords_input):
        self.keywords = [kw.strip() for kw in keywords_input.split(
            ",") if kw.strip()] if keywords_input else []

        # Update the requirements tree
        self.optional_requirements['Entered keyword(s)'] = bool(self.keywords)
        self.populate_tree()

    def on_treeview_select(self, _):
        if self.current_standards_tree:
            item = self.current_standards_tree.selection()[0]
            standard = self.current_standards_tree.item(item)["values"]
            self.selected_standard = standard
            # print(f"Selected standard: {self.selected_standard}")

            if self.selected_current_file and self.selected_new_file:
                self.compare_button.config(state="normal")

        # Update the currently selected standard label
        if self.currently_selected_label:
            self.currently_selected_label.destroy()

        self.currently_selected_label = tk.Label(
            self.currently_selected_frame, textvariable=self.currently_selected_label_text,
            font=("Calibri", 11))
        self.currently_selected_label.pack()

        if self.selected_standard:
            self.currently_selected_label_text.set(
                f"Currently selected standard: {self.selected_standard[0]}")

        # Update the requirements tree
        self.obligatory_requirements['Chosen a standard to compare'] = True
        self.populate_tree()

    def on_matching_treeview_select(self, _):
        if self.matching_new_standards_tree:
            item = self.matching_new_standards_tree.selection()[0]
            standard = self.matching_new_standards_tree.item(item)["values"]
            self.selected_new_standard = standard
            # print(f"\nSelected new standard: {self.selected_new_standard}")
            # print(f"Selected standard: {self.selected_standard}\n")

        if self.write_selected_new_standard_button:
            self.write_selected_new_standard_button.config(state="normal")

    def start_comparison(self):  # pylint: disable=R0915
        if self.selected_standard:
            curr_std = self.selected_standard[0]
            self.matches = {self.selected_standard[0]: {}}
            original_standard = next(
                (std for std in self.current_standards if std["id"] == self.selected_standard[0]), None)
        else:
            curr_std = "Unknown"
            self.matches = {}
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

                        self.matches[curr_std][new_standard["id"]] = {
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

                        self.matches[curr_std][new_standard["id"]] = {
                            "weighted_similarity": round(100 * weighted_similarity, 2),
                            "cosine": cosine_sim,
                            "edit": edit_dist,
                        }

        show_matches(original_standard, self.matches, curr_std)

        # Add the matching new standards to self.potential_new_standards
        self.potential_new_standards = []
        sorted_new_standards = sorted(self.new_standards, key=lambda x: self.matches[curr_std].get(
            x["id"], {}).get("weighted_similarity", 0), reverse=True)
        # print(f"Sorted new standards: {len(sorted_new_standards)}")
        # print(f"First sorted new standard: {sorted_new_standards[0]}")
        for new_standard in sorted_new_standards:
            if new_standard["id"] in self.matches[curr_std]:
                self.potential_new_standards.append(new_standard)
        # print(f"Potential new standards: {len(self.potential_new_standards)}")
        # print(f"First potential new standard: {self.potential_new_standards[0]}")

        if self.process_next_std_button:
            self.process_next_std_button.destroy()

        self.process_next_std_button = tk.Button(
            self.master, text="Process Next Standard (>>>)", command=self.process_next_standard,
            width=114, font=("Calibri", 11))
        self.process_next_std_button.pack(pady=5)
        self.process_next_std_button.config(bg="#b3ffb3")
        self.process_next_std_button.bind(
            "<Enter>", lambda event,
            button=self.process_next_std_button: self.change_cursor(event, button))
        self.process_next_std_button.bind(
            "<Leave>", lambda _: self.master.config(cursor=""))

        if len(self.filtered_standards) == 1:
            self.process_next_std_button.config(state="disabled")

        # Display the matching new standards in a Treeview
        if self.matching_new_standards_tree:
            self.matching_new_standards_tree.destroy()

        # Destroy the previous frame and scrollbar if they exist
        if self.more_matches_frame:
            self.more_matches_frame.destroy()

        if self.new_stds_frame:
            self.new_stds_frame.destroy()

        if self.new_stds_scrollbar:
            self.new_stds_scrollbar.destroy()

        # Entry to allow filtering of new standards
        if self.filter_new_stds_entry:
            self.filter_new_stds_entry.destroy()

        self.filter_new_stds_entry = tk.Entry(
            self.master, width=114, fg="gray", font=("Calibri", 11))
        self.filter_new_stds_entry.pack(pady=5)
        self.filter_new_stds_entry.bind("<KeyRelease>", lambda _: self.filter_new_standards(
            self.filter_new_stds_entry.get()))
        self.add_placeholder(self.filter_new_stds_entry,
                             "Enter the name of the criteria or its No.")

        self.new_stds_frame = tk.Frame(self.master)
        self.new_stds_frame.pack(pady=10)

        self.matching_new_standards_tree = ttk.Treeview(
            self.new_stds_frame, columns=("No.", "New Criteria", "Level", "Similarity"), show="headings")
        self.matching_new_standards_tree.heading("No.", text="No.")
        self.matching_new_standards_tree.heading(
            "New Criteria", text="New Criteria")
        self.matching_new_standards_tree.heading("Level", text="Level")
        self.matching_new_standards_tree.heading(
            "Similarity", text="Similarity")
        self.matching_new_standards_tree.column("#0", width=0, stretch=tk.NO)
        self.matching_new_standards_tree.column("No.", width=50, stretch=tk.NO)
        self.matching_new_standards_tree.column(
            "New Criteria", width=650, stretch=tk.NO)
        self.matching_new_standards_tree.column(
            "Level", width=100, stretch=tk.NO)
        self.matching_new_standards_tree.column(
            "Similarity", width=100, stretch=tk.NO)
        self.matching_new_standards_tree.pack(
            in_=self.new_stds_frame, side="left", fill="both", expand=True)

        self.new_stds_scrollbar = ttk.Scrollbar(
            self.new_stds_frame, orient="vertical", command=self.matching_new_standards_tree.yview)
        self.matching_new_standards_tree.configure(
            yscrollcommand=self.new_stds_scrollbar.set)
        self.new_stds_scrollbar.pack(side="right", fill="y")

        self.matching_new_standards_tree.bind(
            "<ButtonRelease-1>", self.on_matching_treeview_select)

        # Initially, show only the top 10 matches
        for new_standard in self.potential_new_standards[:10]:
            self.matching_new_standards_tree.insert("", "end", values=(
                new_standard["id"], new_standard["text"], new_standard["level"],
                self.matches[curr_std].get(new_standard["id"], {}).get("weighted_similarity", 0)))

        # Add two buttons side by side to show +10 more matches or to show all matches
        if self.more_matches_frame:
            self.more_matches_frame.destroy()

        self.more_matches_frame = tk.Frame(self.master)
        self.more_matches_frame.pack(pady=5)

        if self.show_more_matches_button:
            self.show_more_matches_button.destroy()

        self.show_more_matches_button = tk.Button(
            self.more_matches_frame, text="Show +10 More Matches",
            command=self.show_more_matches, width=30, font=("Calibri", 11))
        self.show_more_matches_button.pack(side=tk.LEFT, padx=5)
        self.show_more_matches_button.config(bg="#b3ffb3")
        self.show_more_matches_button.bind(
            "<Enter>", lambda event,
            button=self.show_more_matches_button: self.change_cursor(event, button))
        self.show_more_matches_button.bind(
            "<Leave>", lambda _: self.master.config(cursor=""))

        if self.show_all_matches_button:
            self.show_all_matches_button.destroy()

        self.show_all_matches_button = tk.Button(
            self.more_matches_frame, text="Show All Matches",
            command=self.show_all_matches, width=30, font=("Calibri", 11))
        self.show_all_matches_button.pack(side=tk.LEFT, padx=5)
        self.show_all_matches_button.config(bg="#b3e6ff")
        self.show_all_matches_button.bind(
            "<Enter>", lambda event,
            button=self.show_all_matches_button: self.change_cursor(event, button))
        self.show_all_matches_button.bind(
            "<Leave>", lambda _: self.master.config(cursor=""))

        if self.write_selected_new_standard_button:
            self.write_selected_new_standard_button.destroy()

        # Button that writes the selected_new_standard to the Excel file
        self.write_selected_new_standard_button = tk.Button(
            self.master, text="Write Selected New Standard",
            command=self.write_selected_new_standard, width=63, font=("Calibri", 11))
        self.write_selected_new_standard_button.pack(pady=5)
        self.write_selected_new_standard_button.config(bg="#ff9999")
        self.write_selected_new_standard_button.bind(
            "<Enter>", lambda event,
            button=self.write_selected_new_standard_button: self.change_cursor(event, button))
        self.write_selected_new_standard_button.bind(
            "<Leave>", lambda _: self.master.config(cursor=""))
        self.write_selected_new_standard_button.config(state="disabled")

        if self.reset_button:
            self.reset_button.destroy()

        # Button that resets the app without unloading the files
        self.reset_button = tk.Button(
            self.master, text="Reset Comparisons", command=self.reset_state, width=63, font=("Calibri", 11))
        self.reset_button.pack(pady=5)
        self.reset_button.config(bg="#ff9999")
        self.reset_button.bind(
            "<Enter>", lambda event,
            button=self.reset_button: self.change_cursor(event, button))
        self.reset_button.bind(
            "<Leave>", lambda _: self.master.config(cursor=""))

    def write_selected_new_standard(self):
        if self.selected_new_standard:
            standard = self.selected_new_standard
            # Update the Excel file with the selected new standard
            if self.selected_standard and self.current_worksheet and standard:
                success = update_standards(
                    self.selected_standard[0], standard[0], standard[1], standard[2], self.current_worksheet)
                if success:
                    self.show_popup(
                        "Success", "The Excel file has been updated successfully.")
                else:
                    self.show_error_popup(
                        "Error", "An error occurred while updating the Excel file.\nPlease make sure the file is not open.")

    def reset_state(self):
        # Clear the content of the ID/Name Entry
        if self.filter_entry:
            self.filter_entry.delete(0, tk.END)
            self.filter_entry.insert(
                0, "Enter the name of the criteria or its No.")
            self.filter_entry.config(fg="gray")

        # Clear the content of the Keywords Entry
        if self.keywords_entry:
            self.keywords_entry.delete(0, tk.END)
            self.keywords_entry.insert(0, "Enter keywords separated by commas")
            self.keywords_entry.config(fg="gray")

        # Reset the checkboxes
        for checkbox in self.worksheet_radio_buttons:
            checkbox.config(state="normal")

        self.selected_worksheet_index.set(0)
        self.current_worksheet = self.worksheets[0]
        self.current_file_button.config(state="disabled")
        self.selected_current_file = True

        self.display_current_standards()

        self.selected_standard = None

        # Update the requirements tree
        self.obligatory_requirements['Chosen a standard to compare'] = False
        self.optional_requirements['Entered criteria name/No.'] = False
        self.optional_requirements['Entered keyword(s)'] = False
        self.populate_tree()

        # Disable the start comparison button
        if self.compare_button:
            self.compare_button.config(state="disabled")

        # Remove the elements that were created during the comparison
        if self.process_next_std_button:
            self.process_next_std_button.destroy()
            self.process_next_std_button = None

        if self.matching_new_standards_tree:
            self.matching_new_standards_tree.destroy()

        if self.more_matches_frame:
            self.more_matches_frame.destroy()

        if self.new_stds_frame:
            self.new_stds_frame.destroy()

        if self.new_stds_scrollbar:
            self.new_stds_scrollbar.destroy()

        if self.filter_new_stds_entry:
            self.filter_new_stds_entry.destroy()

        if self.write_selected_new_standard_button:
            self.write_selected_new_standard_button.destroy()

        if self.reset_button:
            self.reset_button.destroy()

    def show_more_matches(self):
        if self.matching_new_standards_tree and self.show_more_matches_button:
            for new_standard in self.potential_new_standards[len(self.matching_new_standards_tree.get_children()):len(self.matching_new_standards_tree.get_children()) + 10]:
                if self.selected_standard:
                    self.matching_new_standards_tree.insert("", "end", values=(
                        new_standard["id"], new_standard["text"], new_standard["level"],
                        self.matches[self.selected_standard[0]].get(
                            new_standard["id"], {}).get("weighted_similarity", 0)))

            if len(self.matching_new_standards_tree.get_children()) == len(self.potential_new_standards):
                self.show_more_matches_button.config(state="disabled")

    def show_all_matches(self):
        if self.matching_new_standards_tree and self.show_all_matches_button and self.show_more_matches_button:
            for new_standard in self.potential_new_standards:
                if self.selected_standard:
                    self.matching_new_standards_tree.insert("", "end", values=(
                        new_standard["id"], new_standard["text"], new_standard["level"],
                        self.matches[self.selected_standard[0]].get(
                            new_standard["id"], {}).get("weighted_similarity", 0)))

            self.show_more_matches_button.config(state="disabled")
            self.show_all_matches_button.config(state="disabled")

    def show_popup(self, title, message):
        messagebox.showinfo(title, message)

    def show_error_popup(self, title, message):
        messagebox.showerror(title, message)


root = tk.Tk()
root.state("zoomed")
app = StandardsHelperApp(root)
root.mainloop()
