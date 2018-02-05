import autoroster.core

def test_can_open_roster():
    testr = autoroster.core.open_roster("./testrosters/testrosterblank.xlsx")

def test_roster_is_data_frame():
    testr = autoroster.core.open_roster("./testrosters/testrosterblank.xlsx")
    assert str(type(testr)) == "<class 'pandas.core.frame.DataFrame'>"

def test_roster_contains_proper_header():
    testr = autoroster.core.open_roster("./testrosters/testrosterblank.xlsx")
    assert "Student Name" in list(testr)
