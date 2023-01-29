# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionCondFormat20 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_cond_format20
    @xlsx = 'cond_format20.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 10)
    worksheet.write('A2', 20)
    worksheet.write('A3', 30)
    worksheet.write('A4', 40)

    worksheet.conditional_formatting(
      'A1:A4',
      {
        type:       'icon_set',
        icon_style: '3_arrows',
        icons:      [
          { criteria: '>',  type: 'percent', value: 0 },
          { criteria: '<',  type: 'percent', value: 0 },
          { criteria: '>=', type: 'percent', value: 0 }
        ]
      }
    )

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
