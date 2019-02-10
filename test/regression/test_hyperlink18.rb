# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink18 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink18
    @xlsx = 'hyperlink18.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Test long URL at Excel limit.
    worksheet.write_url('A1', 'http://google.com/00000000001111111111222222222233333333334444444444555555555566666666666777777777778888888888999999999990000000000111111111122222222223333333333444444444455555555556666666666677777777777888888888899999999999000000000011111111112222222222x')

    workbook.close
    compare_for_regression(
                                {},
                                { 'xl/workbook.xml' => ['<workbookView'] }
                                )
  end
end
