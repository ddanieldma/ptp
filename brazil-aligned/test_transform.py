# test_transformations.py

import pytest
import pandas as pd
from transform import get_ws, get_color_index, categorize_rows, transform_dataset

@pytest.fixture(scope="module")
def ws():
    """Fixture to load worksheet once per test session."""
    return get_ws()

@pytest.fixture(scope="module")
def colors_dict(ws):
    """Fixture to provide colors used for categorization."""
    return {
        "branco": get_color_index(ws["D1"]),
        "amarelo": get_color_index(ws["D16"]),
        "vermelho": get_color_index(ws["D257"]),
    }

def test_get_color_index(ws):
    """Test color extraction from cells."""
    assert get_color_index(ws["D1"]) == '00000000', "D1 should be white"
    assert get_color_index(ws["D16"]) == 'FFFFFF00', "D16 should be yellow"
    assert get_color_index(ws["D257"]) == 'FFFF0000', "D257 should be red"

def test_categorize_rows(ws, colors_dict):
    """Test row categorization based on cell colors."""
    categories = categorize_rows(ws, 'D', colors_dict)
    assert categories[0] == "neutra", "First row should be 'neutra'"
    assert categories[15] == "contra", "Row 16 should be 'contra'"
    assert categories[256] == "alinhada", "Row 257 should be 'alinhada'"

@pytest.fixture(scope="module")
def sample_dataframe():
    """Sample DataFrame for transformation tests."""
    data = {
        'President': ['A', None, 'B'],
        'Year': [2020, None, 2021],
        'Party': ['A', 'B', 'C'],
        'Agency': ['X', 'Y', 'Z'],
        '% Concedico e Parcialmente': [50, 60, 70],
    }
    return pd.DataFrame(data)

def test_transform_dataset(sample_dataframe):
    """Test dataset transformation with categories."""
    categories = ['header', 'contra', 'alinhada', 'neutra']  # Mock categories
    transformed_df = transform_dataset(sample_dataframe, categories)
    
    assert transformed_df.iloc[0]['category'] == 'contra', "Row 0 category mismatch"
    assert transformed_df.iloc[1]['category'] == 'alinhada', "Row 1 category mismatch"
    assert transformed_df.iloc[2]['category'] == 'neutra', "Row 2 category mismatch"
    assert transformed_df.iloc[1]['year'] == 2020, "Year should be forward filled"
    assert transformed_df.iloc[1]['president'] == 'A', "President should be forward-filled"
    assert 'conc_parc' in transformed_df.columns, "Column should be renamed"
    assert not 'Party' in transformed_df.columns, "Column should be removed"
