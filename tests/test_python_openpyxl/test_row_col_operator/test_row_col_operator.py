import os
import sys
sys.path.append(os.path.dirname(os.path.abspath(__file__)) + '\\..\\..\\..\\')
import pytest
import openpyxl
from extern.python_openpyxl.my_class.row_col_operator.row_col_operator import RowColOperator

"""Verify that the class :class:`row_col_operator` to behave properly."""

@pytest.mark.order(1)
def test_end_row():
  """Verify to get end row of target column correctly. 
  Set various values of column in :obj:`test_sheet1` of :obj:`test_tool_openpyxl.xlsx`.
  """
  row_col_operator = RowColOperator()
  row_col_operator.set_ws = openpyxl.load_workbook('../test_tool_openpyxl.xlsx')['test_sheet1']
  row_col_operator.set_target_col = 1
  assert row_col_operator.get_end_row() == 1
  row_col_operator.set_target_col = 2
  assert row_col_operator.get_end_row() == 15
  row_col_operator.set_target_col = 3
  assert row_col_operator.get_end_row() == 20
  row_col_operator.set_target_col = 4
  assert row_col_operator.get_end_row() == 25
  row_col_operator.set_target_col = 5
  assert row_col_operator.get_end_row() == 15
  row_col_operator.set_target_col = 6
  assert row_col_operator.get_end_row() == 3

@pytest.mark.order(2)
def test_end_col():
  """Verify to get end column of target row correctly. 
  Set various values of row in :obj:`test_sheet1` of :obj:`test_tool_openpyxl.xlsx`.
  """
  row_col_operator = RowColOperator()
  row_col_operator.set_ws = openpyxl.load_workbook('../test_tool_openpyxl.xlsx')['test_sheet1']
  row_col_operator.set_target_row = 1
  assert row_col_operator.get_end_col() == 1
  row_col_operator.set_target_row = 3
  assert row_col_operator.get_end_col() == 6
  row_col_operator.set_target_row = 15
  assert row_col_operator.get_end_col() == 5
  row_col_operator.set_target_row = 16
  assert row_col_operator.get_end_col() == 4
  row_col_operator.set_target_row = 21
  assert row_col_operator.get_end_col() == 4
  row_col_operator.set_target_row = 25
  assert row_col_operator.get_end_col() == 4
  row_col_operator.set_target_row = 26
  assert row_col_operator.get_end_col() == 1

if __name__ == "__main__":
  pass