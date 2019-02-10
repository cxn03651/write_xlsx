# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionEscapes04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_escapes04
    @xlsx = 'escapes04.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet

    worksheet.write_url('A1', 'http://www.perl.com/?a=1&b=2')

    workbook.close
    compare_for_regression(
      nil,
      {
        'xl/workbook.xml' => ['<workbookView']
      }
    )
  end
end
