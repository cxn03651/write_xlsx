# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionProperties05 < Minitest::Test
  def setup
    setup_dir_var
    @long_string = 'This is a long string. ' * 11 + 'AA'
  end

  def teardown
    @tempfile.close(true)
  end

  def test_properties05
    @xlsx = 'properties05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    workbook.set_custom_property('Location', 'Café')

    worksheet.set_column('A:A', 70 )
    worksheet.write(
      'A1',
      "Select 'Office Button -> Prepare -> Properties' to see the file properties."
    )

    workbook.close
    compare_for_regression
  end
end
