# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink20 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_hyperlink20
    @xlsx = 'hyperlink20.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    workbook.instance_variable_set('@custom_colors', ['FF0000FF'])

    worksheet = workbook.add_worksheet
    format1   = workbook.add_format(:color => 'blue', :underline => 1)
    format2   = workbook.add_format(:color => 'red',  :underline => 1)

    worksheet.write_url('A1', 'http://www.python.org/1', format1)
    worksheet.write_url('A2', 'http://www.python.org/2', format2)

    workbook.close

    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
