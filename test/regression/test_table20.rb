# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTable20 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_table20
    @xlsx = 'table01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Set the column width to match the taget worksheet.
    worksheet.set_column('C:F', 10.288)

    # Add the table.
    worksheet.add_table('C3:F13')

    e = assert_raises RuntimeError do
      worksheet.add_table(
        'C3:F7',
        :columns => [{ :header => 'Column1' }, { :header => 'column1' }]
      )
    end

    assert_match(/add_table\(\) contains duplicate name:/, e.message)

    workbook.close
  end
end
