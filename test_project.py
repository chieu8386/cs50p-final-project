import pytest
import pandas as pd
from unittest.mock import patch, MagicMock
from project import get_category, get_dealer, filter_columns, load_mapping

# TEST get_category

def test_category_color_laser():
    assert get_category("84E001") == "CL HW"

def test_category_mono_laser():
    assert get_category("84U001") == "ML HW"

def test_category_inkjet():
    assert get_category("84H001") == "INK HW"

def test_category_ink_consumable():
    assert get_category("8ZC001") == "INK CONS"

def test_category_ml_consumable():
    assert get_category("84XX01") == "ML CONS"

def test_category_cl_consumable():
    assert get_category("84GT01") == "CL CONS"

def test_category_scanner():
    assert get_category("5WDE01") == "SCANNER"

def test_category_others():
    assert get_category("XXXXX") == "OTHERS"

def test_category_nan():
    assert get_category("nan") == "OTHERS"

def test_category_empty():
    assert get_category("") == "OTHERS"

# TEST get_dealer

SAMPLE_DICT = {
    "ADG11": "ADG-HCM",
    "ADG12": "ADG-HN",
    "ASD05": "FPT",
    "VS001": "VIEN SON",
}

def test_dealer_found():
    assert get_dealer("CONG TY ADG11", SAMPLE_DICT) == "ADG-HCM"

def test_dealer_fpt():
    assert get_dealer("TEN CONG TY ASD05", SAMPLE_DICT) == "FPT"

def test_dealer_case_insensitive():
    assert get_dealer("cong ty adg12", SAMPLE_DICT) == "ADG-HN"

def test_dealer_not_found():
    
    assert get_dealer("UNKNOWN COMPANY", SAMPLE_DICT) == "UNKNOWN COMPANY"

def test_dealer_empty_dict():
  
    assert get_dealer("CONG TY ADG11", {}) == "CONG TY ADG11"

# TEST filter_columns

def test_filter_columns_keeps_category():
    
    cols = [f"Col{i}" for i in range(50)]
    df = pd.DataFrame(columns=cols)
    df["Category"] = []
    df["Dealer Name"] = []

    result = filter_columns(df)
    assert "Category" in result.columns

def test_filter_columns_keeps_dealer():
    cols = [f"Col{i}" for i in range(50)]
    df = pd.DataFrame(columns=cols)
    df["Category"] = []
    df["Dealer Name"] = []

    result = filter_columns(df)
    assert "Dealer Name" in result.columns

# TEST load_mapping — dùng mock

def test_load_mapping_success():
    
    mock_df = pd.DataFrame({
        "Code": ["ADG11", "ADG12"],
        "Dealer name": ["ADG-HCM", "ADG-HN"]
    })
    with patch("pandas.read_excel", return_value=mock_df):
        result = load_mapping()
        assert result["ADG11"] == "ADG-HCM"
        assert result["ADG12"] == "ADG-HN"

def test_load_mapping_file_not_found():
    
    with patch("pandas.read_excel", side_effect=FileNotFoundError):
        result = load_mapping()
        assert result == {}