# -*- coding: utf-8 -*-
require 'helper'

class TestDateExamples01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_date_examples01
    @xlsx = 'date_examples01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_column('A:A', 30)

    number = 41333.5

    worksheet.write('A1', number)              #   413333.5

    format2 = workbook.add_format(:num_format => 'dd/mm/yy')
    worksheet.write('A2', number, format2)    #   28/02/13

    format3 = workbook.add_format(:num_format => 'mm/dd/yy')
    worksheet.write('A3', number, format3)    #   02/28/13

    format4 = workbook.add_format(:num_format => 'd\\-m\\-yyyy')
    worksheet.write('A4', number, format4)    #   28-2-2013

    format5 = workbook.add_format(:num_format => 'dd/mm/yy\\ hh:mm')
    worksheet.write('A5', number, format5)    #   28/02/13 12:00

    format6 = workbook.add_format(:num_format => 'd\\ mmm\\ yyyy')
    worksheet.write('A6', number, format6)    #   28 Feb 2013

    format7 = workbook.add_format(:num_format => 'mmm\\ d\\ yyyy\\ hh:mm\\ AM/PM')
    worksheet.write('A7', number, format7)    #   Feb 28 2013 12:00 PM

    workbook.close
    compare_for_regression(
                 ['xl/calcChain.xml', '\[Content_Types\].xml', 'xl/_rels/workbook.xml.rels'],
                 {}
                 )
  end
end
