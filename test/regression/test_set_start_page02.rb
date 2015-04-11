# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSetStartPage02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_set_start_page02
    @xlsx = 'set_start_page02.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.start_page = 2
    worksheet.paper = 9

    worksheet.vertical_dpi = 200

    worksheet.write('A1', 'Foo')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx,
                                [],
                                {
                                  'xl/worksheets/sheet1.xml' => ['<pageMargins']
                                }
                                )
  end
end
