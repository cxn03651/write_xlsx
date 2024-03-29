# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink06 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink06
    @xlsx = 'hyperlink06.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    worksheet.write_url('A1',  'external:C:\Temp\foo.xlsx')
    worksheet.write_url('A3',  'external:C:\Temp\foo.xlsx#Sheet1!A1')
    worksheet.write_url('A5',  'external:C:\Temp\foo.xlsx#Sheet1!A1', 'External', nil, 'Tip')

    workbook.close
    compare_for_regression
  end
end
