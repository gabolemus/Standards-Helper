"""Module with functions for working with text."""


def check_keywords(text, present_keywords):
    """Checks if any provided keyword is present in the text."""
    return any(kw in text for kw in present_keywords)
