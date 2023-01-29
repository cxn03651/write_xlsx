# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHeaderImage21 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header_image21
    @xlsx = 'header_image21.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet

    worksheet1.set_portrait
    worksheet2.set_portrait
    worksheet3.set_portrait

    worksheet1.instance_variable_set(:@vertical_dpi, 200)
    worksheet2.instance_variable_set(:@vertical_dpi, 200)
    worksheet3.instance_variable_set(:@vertical_dpi, 200)

    worksheet3.set_header(
      '&L&G',
      nil,
      {
        image_left: 'test/regression/images/red.jpg'
      }
    )

    workbook.close
    compare_for_regression
  end
end
