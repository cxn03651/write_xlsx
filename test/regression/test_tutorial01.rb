# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTutorial01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_tutorial01
    @xlsx = 'tutorial01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    expenses = [
                ['Rent', 1000],
                ['Gas', 100],
                ['Food', 300],
                ['Gym', 50]
               ]
    expenses.each_with_index do |item, index|
      worksheet.write(index, 0, item[0])
      worksheet.write(index, 1, item[1])
    end

    worksheet.write(expenses.size, 0, 'Total')
    worksheet.write(expenses.size, 1, '=SUM(B1:B4)', nil, 1450)

    workbook.close
    compare_for_regression(
                                ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels'],
                                {}
                                )
  end
end
