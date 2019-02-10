# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionAutofilter02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_autofilter02
    @xlsx = 'autofilter02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    data = []
    data_lines.split(/\n/).each { |line| data << line.split }

    worksheet.write('A1', headings)
    worksheet.autofilter(0, 0, 50, 3)
    worksheet.filter_column(0, 'Region eq East')

    # Hide the rows that don't match the filter criteria.
    row = 1
    data.each do |row_data|
      region = row_data[0]
      if region == 'East'
        # Row is visible.
      else
        # Hide row.
        worksheet.set_row(row, nil, nil, 1)
      end
      worksheet.write(row, 0, row_data)
      row += 1
    end

    workbook.close
    compare_for_regression(
                 nil,
                 {'xl/workbook.xml' => ['<workbookView']}
                 )
  end

  def headings
    %w[Region Item Volume Month]
  end

  def data_lines
    <<EOS
East      Apple     9000      July
East      Apple     5000      July
South     Orange    9000      September
North     Apple     2000      November
West      Apple     9000      November
South     Pear      7000      October
North     Pear      9000      August
West      Orange    1000      December
West      Grape     1000      November
South     Pear      10000     April
West      Grape     6000      January
South     Orange    3000      May
North     Apple     3000      December
South     Apple     7000      February
West      Grape     1000      December
East      Grape     8000      February
South     Grape     10000     June
West      Pear      7000      December
South     Apple     2000      October
East      Grape     7000      December
North     Grape     6000      April
East      Pear      8000      February
North     Apple     7000      August
North     Orange    7000      July
North     Apple     6000      June
South     Grape     8000      September
West      Apple     3000      October
South     Orange    10000     November
West      Grape     4000      July
North     Orange    5000      August
East      Orange    1000      November
East      Orange    4000      October
North     Grape     5000      August
East      Apple     1000      December
South     Apple     10000     March
East      Grape     7000      October
West      Grape     1000      September
East      Grape     10000     October
South     Orange    8000      March
North     Apple     4000      July
South     Orange    5000      July
West      Apple     4000      June
East      Apple     5000      April
North     Pear      3000      August
East      Grape     9000      November
North     Orange    8000      October
East      Apple     10000     June
South     Pear      1000      December
North     Grape     10000     July
East      Grape     6000      February
EOS
  end
end
