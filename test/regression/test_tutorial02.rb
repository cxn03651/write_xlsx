# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTutorial02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_tutorial02
    @xlsx = 'tutorial02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    bold      = workbook.add_format(:bold => 1)
    money     = workbook.add_format(:num_format => '\\$#,##0')

    worksheet.write('A1', 'Item', bold)
    worksheet.write('B1', 'Cost', bold)

    expenses = [
                ['Rent', 1000],
                ['Gas', 100],
                ['Food', 300],
                ['Gym', 50]
               ]
    expenses.each_with_index do |item, index|
      worksheet.write(index + 1, 0, item[0])
      worksheet.write(index + 1, 1, item[1], money)
    end

    worksheet.write(expenses.size + 1, 0, 'Total', bold)
    worksheet.write(expenses.size + 1, 1, '=SUM(B2:B5)', money, 1450)

    workbook.close
    compare_for_regression(
                                ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels'],
                                {}
                                )
  end
end
