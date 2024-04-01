# Function to calculate the maximum similarity score for multiple lines of text
from typing import Tuple

from standards.text_comparison import (cosine_similarity_compare,
                                       edit_distance_compare)


def calculate_max_similarity(original_text, new_text):
    max_similarity = 0.0
    for line in new_text.split("\n"):
        cosine_similarity = cosine_similarity_compare(original_text, line)
        edit_distance = edit_distance_compare(original_text, line)
        max_similarity = max(max_similarity, cosine_similarity, edit_distance)
    return max_similarity


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
