"""File with the configurations to run the project."""

# original_standards_file = "/mnt/c/Users/Gabo/Downloads/Comparison STD Gabo.xlsx"
# new_standards_file = "/mnt/c/Users/Gabo/Downloads/DraftUnifiedStandard_PlanningVersion1.0 Gabo.xlsx"
original_standards_file = "C:\\Users\\Gabo\\Downloads\\Comparison STD Gabo.xlsx"
new_standards_file = "C:\\Users\\Gabo\\Downloads\\DraftUnifiedStandard_PlanningVersion1.0 Gabo.xlsx"

original_standards_ws = ["RCS", "GRS", "RWS", "RMS", "RAS", "RDS"]


def get_worksheet_index(worksheet_name: str) -> int:
    """Get the index of the worksheet."""
    return original_standards_ws.index(worksheet_name)
