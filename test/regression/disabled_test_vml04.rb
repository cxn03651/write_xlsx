# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionVml04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_vml04
    @xlsx = 'vml04.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet

    (0..127).each do |row|
      (0..15).each do |col|
        worksheet1.write_comment(row, col, 'Some text')
      end
    end

    worksheet3.write_comment('A1',  'More text')

    # Set the author to match the target XLSX file.
    worksheet1.comments_author = 'John'
    worksheet3.comments_author = 'John'

    worksheet1.insert_button('B2', {})
    worksheet1.insert_button('C4', {})
    worksheet1.insert_button('E6', {})

    worksheet3.insert_button('E8', {})

    workbook.close
    compare_for_regression
  end
end
