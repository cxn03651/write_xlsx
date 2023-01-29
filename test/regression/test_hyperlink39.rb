# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink39 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink39
    @xlsx = 'hyperlink39.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      url: 'external:c:\temp\foo.xlsx'
    )

    workbook.close

    compare_for_regression
  end
end
