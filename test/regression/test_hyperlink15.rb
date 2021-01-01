# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink15 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink15
    @xlsx = 'hyperlink16.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    worksheet.write_url('B2', 'external:./subdir/blank.xlsx')

    workbook.close
    compare_for_regression(
                                {},
                                { 'xl/workbook.xml' => ['<workbookView'] }
                                )
  end
end
