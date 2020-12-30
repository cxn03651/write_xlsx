# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionCondFormat16 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_cond_format16
    @xlsx = 'cond_format16.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # We manually set the indices to get the same order as the target file.
    format1 = workbook.add_format(:bg_color => 'red')
    format2 = workbook.add_format(:bg_color => '#92D050')

    worksheet.write('A1', 10)
    worksheet.write('A2', 20)
    worksheet.write('A3', 30)
    worksheet.write('A4', 40)

    options1 = {
      :type         => 'cell',
      :format       => format1,
      :criteria     => 'less than',
      :value        => 5,
      :stop_if_true => false
    }

    worksheet.conditional_formatting('A1', options1)

    options2 = {
      :type         => 'cell',
      :format       => format2,
      :criteria     => 'greater than',
      :value        => 20,
      :stop_if_true => true
    }

    worksheet.conditional_formatting('A1', options2)

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
