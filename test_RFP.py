import unittest
from RFP_3 import *


class RFPTest(unittest.TestCase):

    def setUp(self) -> None:
        self.RFP = RFP()

    def test_empty_cells(self):
        self.RFP.DTT()
        self.assertEqual(ws['A2'].value, 0)

    def test_average(self):
        self.RFP.DTT()
        self.assertEqual(ws['B12'].value, '%2f' % 2)

    def test_min(self):
        self.RFP.DTT()
        self.assertEqual(ws['G12'].value, '%2f' % 0.8888888)

    def test_max(self):
        self.assertEqual(ws['H12'].value, '%2f' % 5)

    @unittest.skip("WIP")
    def tests_string_in_cell(self):
        self.RFP = RFP_Test()
        self.assertEqual(ws['A2'].value, 0)



        