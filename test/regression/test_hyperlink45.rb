# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink45 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink45
    @xlsx = 'hyperlink45.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      :url => %q(external:\\\\Vboxsvr\share\foo bar.xlsx#'Some Sheet'!A1)
    )

    workbook.close

    compare_for_regression
  end
end
