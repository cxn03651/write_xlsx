# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionSetStartPage03 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_set_start_page03
    @xlsx = 'set_start_page03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.start_page = 101
    worksheet.paper = 9

    worksheet.vertical_dpi = 200

    worksheet.write('A1', 'Foo')

    workbook.close
    compare_for_regression(
      [],
      {
        'xl/worksheets/sheet1.xml' => ['<pageMargins']
      }
    )
  end
end
