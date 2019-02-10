# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTutorial03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_tutorial03
    @xlsx = 'tutorial03.xlsx'
    workbook     = WriteXLSX.new(@io)
    worksheet    = workbook.add_worksheet

    bold         = workbook.add_format(:bold => 1)
    money_format = workbook.add_format(:num_format => '\\$#,##0')
    date_format  = workbook.add_format(:num_format => 'mmmm\\ d\\ yyyy')

    worksheet.set_column('B:B', 15)

    worksheet.write('A1', 'Item', bold)
    worksheet.write('B1', 'Date', bold)
    worksheet.write('C1', 'Cost', bold)

    expenses = [
                [ 'Rent', '2013-01-13T', 1000 ],
                [ 'Gas',  '2013-01-14T', 100 ],
                [ 'Food', '2013-01-16T', 300 ],
                [ 'Gym',  '2013-01-20T', 50 ]
               ]
    expenses.each_with_index do |item, index|
      worksheet.write_string(index + 1,    0, item[0])
      worksheet.write_date_time(index + 1, 1, item[1], date_format)
      worksheet.write_number(index + 1,    2, item[2], money_format)
    end

    worksheet.write(expenses.size + 1, 0, 'Total', bold)
    worksheet.write(expenses.size + 1, 2, '=SUM(C2:C5)', money_format, 1450)

    workbook.close
    compare_for_regression(
                                ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels'],
                                {}
                                )
  end
end
