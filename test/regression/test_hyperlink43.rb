# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink43 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink43
    @xlsx = 'hyperlink43.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      url: 'external:c:\te mp\foo.xlsx'
    )

    workbook.close

    compare_for_regression
  end
end
