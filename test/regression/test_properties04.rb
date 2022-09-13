# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionProperties04 < Minitest::Test
  def setup
    setup_dir_var
    @long_string = ('This is a long string. ' * 11) + 'AA'
  end

  def teardown
    @tempfile.close(true)
  end

  def test_properties04
    @xlsx = 'properties04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    workbook.set_custom_property('Checked by',      'Adam',                 'text')
    workbook.set_custom_property('Date completed',  '2016-12-12T23:00:00Z', 'date')
    workbook.set_custom_property('Document number', '12345',               'number_int')
    workbook.set_custom_property('Reference',       '1.2345',               'number')
    workbook.set_custom_property('Source',          true,                   'bool')
    workbook.set_custom_property('Status',          false,                  'bool')
    workbook.set_custom_property('Department',      @long_string,           'text')
    workbook.set_custom_property('Group',           '1.2345678901234',      'number')

    worksheet.set_column('A:A', 70)
    worksheet.write(
      'A1',
      "Select 'Office Button -> Prepare -> Properties' to see the file properties."
    )

    workbook.close
    compare_for_regression
  end

  def test_properties04_2
    @xlsx = 'properties04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    workbook.set_custom_property('Checked by',      'Adam')
    workbook.set_custom_property('Date completed',  '2016-12-12T23:00:00Z', 'date')
    workbook.set_custom_property('Document number', '12345')
    workbook.set_custom_property('Reference',       '1.2345')
    workbook.set_custom_property('Source',          1,                      'bool')
    workbook.set_custom_property('Status',          nil,                    'bool')
    workbook.set_custom_property('Department',      @long_string)
    workbook.set_custom_property('Group',           '1.2345678901234')

    worksheet.set_column('A:A', 70)
    worksheet.write(
      'A1',
      "Select 'Office Button -> Prepare -> Properties' to see the file properties."
    )

    workbook.close
    compare_for_regression
  end
end
