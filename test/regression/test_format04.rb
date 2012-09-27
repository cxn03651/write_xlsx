# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFormat04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_format04
    @xlsx = 'format04.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    format1 = workbook.add_format
    format2 = workbook.add_format(:bold => 1)

    # Test the copy method
    format2.copy(format1)
    format2.set_italic
    format2.set_bold

    worksheet.write('A1', 'Foo', format2)

    worksheet.conditional_formatting(
                                     'C1:C10',
                                     {
                                       :type     => 'cell',
                                       :criteria => '>',
                                       :value    => 50,
                                       :format   => format2
                                     }
                                     )

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
