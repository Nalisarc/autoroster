import autoroster.core
from nose.tools import with_setup

def setup():
    testr = autoroster.core.open_roster("./testrosters/testrosterminimal.xlsx")

@with_setup(setup)
def test_organizes_by_exam():
    test_exams = autoroster.core.get_exam_list()
    assert any(['Biol 3040','Cs 2050','Engl 4050']) in test_exams
