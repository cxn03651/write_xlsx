# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink30 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink30
    @xlsx = 'hyperlink30.xlsx'
    workbook  = WriteXLSX.new(@io)

    # Simulate custom color for testing.
    workbook.instance_variable_set(:@custom_colors, ['FF0000FF'])

    worksheet = workbook.add_worksheet
    format1   = workbook.add_format(hyperlink: 1)
    format2   = workbook.add_format(color: 'red',  underline: 1)
    format3   = workbook.add_format(color: 'blue', underline: 1)

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    worksheet.write_url('A1', 'http://www.python.org/1', format1)
    worksheet.write_url('A2', 'http://www.python.org/2', format2)
    worksheet.write_url('A3', 'http://www.python.org/3', format3)

    workbook.close

    compare_for_regression
  end
end
