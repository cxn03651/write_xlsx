# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionProperties04 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_properties04
    @xlsx = 'properties04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    long_string = 'This is a long string. ' * 11 + 'AA'

    workbook.set_custom_property('Checked by',      'Adam'                            )
    workbook.set_custom_property('Date completed',  '2016-12-12T23:00:00Z', 'date'    )
    workbook.set_custom_property('Document number', '12345' ,               'num_int' )
    workbook.set_custom_property('Reference',       '1.2345',               'num_real')
    workbook.set_custom_property('Source',          1,                      'bool'    )
    workbook.set_custom_property('Status',          0,                      'bool'    )
    workbook.set_custom_property('Department',      long_string                       )
    workbook.set_custom_property('Group',           '1.2345678901234',      'num_real')

    worksheet.set_column('A:A', 70 )
    worksheet.write(
      'A1',
      "Select 'Office Button -> Prepare -> Properties' to see the file properties."
    )

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
