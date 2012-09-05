# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionComment05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_comment05
    @xlsx = 'comment05.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet

    (0..127).each do |row|
      (0..15).each {|col| worksheet1.write_comment(row, col, 'Some text')}
    end

    worksheet3.write_comment('A1', 'More text')

    # Set the author to match the target XLSX file.
    worksheet1.set_comments_author('John')
    worksheet3.set_comments_author('John')

    workbook.close
    compare_xlsx_for_regression(
                                File.join(@regression_output, @xlsx),
                                @xlsx,
                                nil,
                                { 'xl/workbook.xml' => ['<workbookView'] })
  end
end
