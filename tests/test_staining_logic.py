import pandas as pd

from staining_logic import parse_dilution_ratio, validate_and_prepare_df


def test_parse_dilution_ratio_success():
    assert parse_dilution_ratio("1:100") == 100


def test_parse_dilution_ratio_invalid():
    try:
        parse_dilution_ratio("100")
        assert False, "应当抛出 ValueError"
    except ValueError:
        assert True


def test_validate_and_prepare_df_basic():
    df = pd.DataFrame(
        {
            "marker": ["CD3"],
            "荧光染料": ["FITC"],
            "稀释比例": ["1:100"],
            "是否作为FMO": ["是"],
            "一抗/二抗/胞内抗体": ["胞内"],
        }
    )

    prepared, errors = validate_and_prepare_df(df)
    assert errors == []
    assert prepared.iloc[0]["抗体类型"] == "胞内抗体"
    assert bool(prepared.iloc[0]["是否作为FMO"]) is True


def test_secondary_antibody_fmo_should_be_ignored():
    df = pd.DataFrame(
        {
            "marker": ["CD19"],
            "荧光染料": ["PE"],
            "稀释比例": ["1:200"],
            "是否作为FMO": ["是"],
            "一抗/二抗/胞内抗体": ["二抗"],
        }
    )

    prepared, errors = validate_and_prepare_df(df)
    assert errors == []
    assert prepared.iloc[0]["抗体类型"] == "二抗"
    assert bool(prepared.iloc[0]["是否作为FMO"]) is False


def test_autofluorescence_does_not_require_dilution():
    df = pd.DataFrame(
        {
            "marker": ["YFP"],
            "荧光染料": ["YFP"],
            "稀释比例": [""],
            "是否作为FMO": ["是"],
            "一抗/二抗/胞内抗体": ["自发荧光"],
        }
    )

    prepared, errors = validate_and_prepare_df(df)
    assert errors == []
    assert prepared.iloc[0]["抗体类型"] == "自发荧光"
