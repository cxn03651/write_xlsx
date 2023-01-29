# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTypes08 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_types08
    @xlsx = 'types08.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    bold   = workbook.add_format(bold: 1)
    italic = workbook.add_format(italic: 1)

    worksheet.write_boolean('A1', 'True', bold)
    worksheet.write_boolean('A2', nil, italic)

    workbook.close

    compare_for_regression(
      ['xl/styles.xml'],
      {}
    )
  end
end
