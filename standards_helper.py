"""Contains the logic for the program that helps link the standards."""

from standards.text_comparison import (cosine_similarity_compare,
                                       edit_distance_compare)
from standards.workbooks import get_new_standards, get_original_standards

original_standards = get_original_standards()
# print("Original standards:")
# for standard in original_standards:
#     print(f"{standard['id']}: {standard['text']}")
#     # print(f"{standard['id']}: {standard['text']} | level: {standard['level']}")

# Begin reading the new standards file from row 4, columns A, B and C
new_standards = get_new_standards()
# print("\nNew standards:")
# for standard in new_standards:
#     # print(f"{standard['id']}: {standard['text']}")
#     print(f"{standard['id']}: {standard['text']} | level: {standard['level']}")

# Let the user choose the original standard
while True:
    user_option = input(
        "Please enter the number of the standard to compare (enter 'quit' to exit): ")

    if user_option.lower() in ['q', 'quit']:
        break

    matches = {
        user_option: {}
    }
    original_standard = list(filter(
        lambda x: x["id"] == user_option, original_standards))[0]

    for new_standard in new_standards:
        if new_standard["text"]:
            if "\n" in new_standard["text"]:
                scores = []
                cosine_scores = []
                edit_scores = []
                for line in new_standard["text"].split("\n"):
                    cosine_similarity = cosine_similarity_compare(
                        original_standard["text"] or "", line)
                    edit_distance = edit_distance_compare(
                        original_standard["text"] or "", line)
                    cosine_scores.append(cosine_similarity)
                    edit_scores.append(edit_distance)
                    scores.append(max(cosine_similarity, edit_distance))
                max_val_idx = scores.index(max(scores))
                cosine_val = cosine_scores[max_val_idx]
                edit_val = edit_scores[max_val_idx]
                matches[user_option][new_standard["id"]] = {
                    "cosine": cosine_val,
                    "edit": edit_val,
                    "max": max(scores)
                }
            else:
                cosine_similarity = cosine_similarity_compare(
                    original_standard["text"] or "", new_standard["text"] or "")
                edit_distance = edit_distance_compare(
                    original_standard["text"] or "", new_standard["text"] or "")
                matches[user_option][new_standard["id"]] = {
                    "cosine": cosine_similarity,
                    "edit": edit_distance,
                    "max": max(cosine_similarity, edit_distance)
                }

    # Show the 10 new standards with the highest similarity; either cosine or edit distance
    print(f"Original standard: {original_standard['text']}")
    print("Top 10 matches:")
    for i, match in enumerate(sorted(matches[user_option].items(), key=lambda x: x[1]["max"], reverse=True)[:10]):
        print(f"{i+1}. {match[0]}: {match[1]}")
    print("\n")
