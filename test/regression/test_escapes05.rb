# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionEscapes05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_escapes05
    @xlsx = 'escapes05.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet1  = workbook.add_worksheet('Start')
    worksheet2  = workbook.add_worksheet('A & B')

    worksheet1.write_url('A1', "internal:'A & B'!A1", 'Jump to A & B')

    workbook.close
    compare_for_regression(
      nil,
      {
        'xl/workbook.xml' => ['<workbookView']
      }
    )
  end
end
