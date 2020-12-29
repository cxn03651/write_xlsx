# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionProperties03 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_properties03
    @xlsx = 'properties03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    workbook.set_custom_property('Checked by', 'Adam')

    worksheet.set_column('A:A', 70)
    worksheet.write('A1', "Select 'Office Button -> Prepare -> Properties' to see the file properties.")

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
