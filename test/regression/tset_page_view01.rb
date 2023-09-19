# -*- coding: utf-8 -*-

require 'helper'

class TestPageView01 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_page_view01
    @xlsx = 'page_view01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_page_view
    # Options to match automatic page setup.
    worksheet.set_paper(9)
    worksheet.vertical_dpi = 200

    worksheet.write('A1', 'Foo')

    workbook.close
    compare_for_regression
  end
end
