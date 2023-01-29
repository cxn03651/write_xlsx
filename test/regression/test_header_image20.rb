# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHeaderImage20 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header_image20
    @xlsx = 'header_image20.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_header(
      '&C&G',
      nil,
      {
        :image_center => 'test/regression/images/watermark.png'
      }
    )

    workbook.close
    compare_for_regression(
      [],
      {
        'xl/worksheets/sheet1.xml' => ['<pageMargins', '<pageSetup']
      }
    )
  end
end
