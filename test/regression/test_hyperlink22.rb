# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink22 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink22
    @xlsx = 'hyperlink22.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    worksheet.write_url('A1', 'external:\\\\Vboxsvr\share\foo bar.xlsx')

    workbook.close

    compare_for_regression
  end
end
