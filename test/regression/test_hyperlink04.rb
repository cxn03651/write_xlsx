# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink04
    @xlsx = 'hyperlink04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet('Data Sheet')

    worksheet1.write_url('A1',  "internal:Sheet2!A1")
    worksheet1.write_url('A3',  "internal:Sheet2!A1:A5")
    worksheet1.write_url('A5',  "internal:'Data Sheet'!D5", 'Some text')
    worksheet1.write_url('E12', "internal:Sheet1!J1")
    worksheet1.write_url('G17', "internal:Sheet2!A1", 'Some text', nil)
    worksheet1.write_url('A18', "internal:Sheet2!A1", nil, nil, 'Tool Tip 1')
    worksheet1.write_url('A20', "internal:Sheet2!A1", 'More text', nil, 'Tool Tip 2')

    workbook.close
    compare_for_regression
  end
end
