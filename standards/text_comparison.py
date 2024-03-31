from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


def cosine_similarity_compare(text1: str, text2: str) -> float:
    """Compares two different strings using the cosine similarity method."""
    tfidf_vectorizer = TfidfVectorizer()  # Create the TF-IDF vectorizer
    # Fit and transform the texts
    tfidf_matrix = tfidf_vectorizer.fit_transform([text1, text2])
    # Calulate the cosine similarity
    cosine_sim = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])

    return cosine_sim[0][0]
