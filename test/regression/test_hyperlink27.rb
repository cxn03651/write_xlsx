# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink27 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink27
    @xlsx = 'hyperlink27.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    worksheet.write_url('A1', %q(external:\\\\Vboxsvr\share\foo bar.xlsx#'Some Sheet'!A1))

    workbook.close

    compare_for_regression
  end
end
