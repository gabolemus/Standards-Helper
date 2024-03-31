"""Contains the logic for the program that helps link the standards."""

from typing import Any, Tuple

from standards.text_comparison import (cosine_similarity_compare,
                                       edit_distance_compare)
from standards.workbooks import (get_new_standards, get_original_standards,
                                 update_standards)


# Function to calculate the maximum similarity score for multiple lines of text
def calculate_max_similarity(original_text, new_text):
    max_similarity = 0.0
    for line in new_text.split("\n"):
        cosine_similarity = cosine_similarity_compare(original_text, line)
        edit_distance = edit_distance_compare(original_text, line)
        max_similarity = max(max_similarity, cosine_similarity, edit_distance)
    return max_similarity


original_standards = get_original_standards()
new_standards = get_new_standards()


def show_matches(original_standard, matches, user_option, amount=10):
    if original_standard:
        print(f"Original standard: {original_standard['text']}")
    print(f"Top {amount} matches:")
    for i, (std_id, match) in enumerate(sorted(matches[user_option].items(),
                                               key=lambda x: x[1]["weighted_similarity"],
                                               reverse=True)[:amount]):
        print(f"{i+1}. {std_id}: {match}")
    print("\n")


def get_text_comparisons(text1: str, text2: str) -> Tuple[float, float]:
    """Gets the comparisons values between two texts."""
    if "\n" not in text2:
        cosine_sim_val = cosine_similarity_compare(text1, text2)
        edit_dist_val = edit_distance_compare(text1, text2)
        return (cosine_sim_val, edit_dist_val)

    cosine_sims = []
    edit_dists = []
    max_cosine_sim = 0.0
    max_edit_dist = 0.0

    for line in text2.split("\n"):
        cosine_sim_val = cosine_similarity_compare(text1, line)
        edit_dist_val = edit_distance_compare(text1, line)

        max_cosine_sim = max(max_cosine_sim, cosine_sim_val)
        max_edit_dist = max(max_edit_dist, edit_dist_val)

        cosine_sims.append(cosine_sim_val)
        edit_dists.append(edit_dist_val)

    if max_cosine_sim > max_edit_dist:
        return (max_cosine_sim, cosine_sims[cosine_sims.index(max_cosine_sim)])

    return (max_edit_dist, edit_dists[edit_dists.index(max_edit_dist)])


# Let the user choose the original standard
while True:
    matches_to_show = 10

    user_option = input(
        "Please enter the number of the standard to compare (enter 'quit' to exit): ")

    if user_option.lower() in ['e', 'exit', 'q', 'quit']:
        break

    # Ask the user for keywords
    keywords_input = input(
        "Enter keywords separated by commas (or leave blank): ")

    # Process keywords
    keywords = [kw.strip() for kw in keywords_input.split(",")
                if kw.strip()] if keywords_input else []

    matches: Any = {
        user_option: {}
    }
    original_standard = next(
        (std for std in original_standards if std["id"] == user_option), None)

    if original_standard:
        for new_standard in new_standards:
            if new_standard["text"] and original_standard["text"]:
                cosine_sim, edit_dist = get_text_comparisons(
                    original_standard["text"], new_standard["text"])

                if keywords:
                    cosine_weight = 0.3
                    edit_weight = 0.3
                    keyword_weight = 0.4

                    keyword_proportion = len(
                        [kw for kw in keywords if kw in new_standard["text"]]) / len(keywords)
                    weighted_similarity = cosine_sim * cosine_weight + edit_dist * \
                        edit_weight + keyword_proportion * keyword_weight

                    matches[user_option][new_standard["id"]] = {
                        "weighted_similarity": weighted_similarity,
                        "cosine": cosine_sim,
                        "edit": edit_dist,
                        "keyword_proportion": keyword_proportion,
                    }
                else:
                    cosine_weight = 0.5
                    edit_weight = 0.5

                    weighted_similarity = cosine_sim * cosine_weight + \
                        edit_dist * edit_weight

                    matches[user_option][new_standard["id"]] = {
                        "weighted_similarity": weighted_similarity,
                        "cosine": cosine_sim,
                        "edit": edit_dist,
                    }

    show_matches(original_standard, matches, user_option)

    # Ask the user for the next action: show more matches, choose another
    # standard, exit or write the new standards to the current standards file
    next_action = input(
        "Enter 'more' to show more matches, 'choose' to choose another standard, 'exit' to exit, or 'write' to write the new standards to the current standards file: ")

    if next_action.lower() in ['e', 'exit', 'q', 'quit']:
        break

    if next_action.lower() in ['m', 'more']:
        matches_to_show += 10
        show_matches(original_standard, matches, user_option, matches_to_show)

    if next_action.lower() in ['c', 'choose']:
        continue

    if next_action.lower() in ['w', 'write']:
        # Write the new standards to the current standards file
        standard_id = input(
            "Enter the index of the standard to write to the current standards file: ")

        if standard_id in matches[user_option]:
            print("Writing the new standard to the current standards file.")
            match = matches[user_option][standard_id]
            # Filter the new standard to show the one that matches the id
            new_standard = filter(
                lambda x: x["id"] == standard_id, new_standards)
            new_standard = next(new_standard, None)
            if new_standard:
                print(f"New standard: {new_standard}")
                update_standards(user_option, new_standard["id"] or "",
                                 new_standard["text"] or "", new_standard["level"] or "", "RCS")
            continue
