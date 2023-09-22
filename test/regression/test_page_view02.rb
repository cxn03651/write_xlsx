# -*- coding: utf-8 -*-

require 'helper'

class TestPageView02 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_page_view02
    @xlsx = 'page_view02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_page_view
    worksheet.zoom = 75

    # Options to match automatic page setup.
    worksheet.paper = 9
    worksheet.vertical_dpi = 200

    worksheet.write('A1', 'Foo')

    workbook.close
    compare_for_regression
  end
end
