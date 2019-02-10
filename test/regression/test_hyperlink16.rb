# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink16 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink16
    @xlsx = 'hyperlink16.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_url('B2', 'external:./subdir/blank.xlsx')

    workbook.close
    compare_for_regression(
                                {},
                                { 'xl/workbook.xml' => ['<workbookView'] }
                                )
  end
end
