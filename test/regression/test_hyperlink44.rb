# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink44 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink44
    @xlsx = 'hyperlink44.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet('Sheet 1')

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      :url => "internal:'Sheet 1'!A1"
    )

    workbook.close

    compare_for_regression
  end
end
