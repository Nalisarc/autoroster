
import autoroster.core
from nose.tools import with_setup

def setup():
    testr = autoroster.core.open_roster("./testrosters/testrosterminimal.xlsx")

def test_organizes_by_exam():
    testr = autoroster.core.open_roster("./testrosters/testrosterminimal.xlsx")
    test_exams = autoroster.core.get_exam_list(testr)
    assert any([(x in ['Biol 3040', 'Cs 2050', 'Engl 4050']) for x in test_exams]), test_exams
