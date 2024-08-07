import pytest
import sys
import os

# Add the parent directory to the system path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from edusoul import extract_phone_numbers

def test_extract_phone_numbers():
    # Example test case with mock data
    test_text = "Contact me at +1234567890 or +1987654321."
    expected_numbers = ["+1234567890", "+1987654321"]
    extracted_numbers = extract_phone_numbers(test_text)
    assert extracted_numbers == expected_numbers
