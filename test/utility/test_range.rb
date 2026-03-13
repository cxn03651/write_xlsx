# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx/utility/common'
require 'write_xlsx/utility/cell_reference'

class TestRange < Minitest::Test
  include Writexlsx::Utility::Common
  include Writexlsx::Utility::CellReference

  def test_range_0_0_1_1
    assert_equal(
      'B1',
      xl_range(0, 0, 1, 1)
    )
  end

  def test_range_0_0_1_1_1_1_1_1
    assert_equal(
      '$B$1',
      xl_range(0, 0, 1, 1, 1, 1, 1, 1)
    )
  end
end
