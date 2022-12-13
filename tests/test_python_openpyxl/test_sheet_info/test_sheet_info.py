import os
import sys
sys.path.append(os.path.dirname(os.path.abspath(__file__)) + '\\..\\..\\..\\')
import pytest
import openpyxl
from extern.python_openpyxl.my_class.sheet_info.sheet_info import SheetInfo

"""Verify that the class :class:`SheetInfo` to behave properly."""

@pytest.mark.order(1)
def test_set_sheet_info():
  """Verify that the class set value properly. 
  Set values in :obj:`test_sheet1` of :obj:`test_tool_openpyxl.xlsx`.
  Verify get end row and end column properly.
  """
  sheet_info = SheetInfo()
  sheet_info.set_wb_name = '../test_tool_openpyxl.xlsx'
  sheet_info.set_ws_name = 'test_sheet1'
  sheet_info.set_sheet_info()
  sheet_info.set_start_row = 2
  assert sheet_info.get_start_row == 2
  sheet_info.set_start_col = 2
  assert sheet_info.get_start_col == 2
  sheet_info.set_key_row = 4
  assert sheet_info.get_key_row == 4
  sheet_info.set_key_col = 4
  assert sheet_info.get_key_col == 4
  sheet_info.set_target_row = 2
  sheet_info.set_target_col = 5
  sheet_info.set_row_col_info()
  assert sheet_info.get_end_col == 6
  assert sheet_info.get_end_row == 15
  sheet_info.set_result_row = 6
  assert sheet_info.get_result_row == 6
  sheet_info.set_result_col = 6
  assert sheet_info.get_result_col == 6

if __name__ == "__main__":
  pass