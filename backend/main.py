import sys
import os

# ✅ Ensure project root is in PYTHONPATH
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from backend import models
from core import parser, scorer, feedback

def main():
    print("Backend is running...")
    # Example test
    sample_text = "John Doe - Python Developer"
    print("Extracted text:", sample_text)
    print("Scoring:", scorer.score_resume(sample_text))
    print("Feedback:", feedback.generate_feedback(sample_text))

if __name__ == "__main__":
    main()
