# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink07 < Minitest::Test
  def setup
    setup_dir_var
  end

  def test_hyperlink07
    @xlsx = 'hyperlink07.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    worksheet.write_url('A1',  'external:\\\\VBOXSVR\share\foo.xlsx', 'J:\foo.xlsx')
    worksheet.write_url('A3',  'external:foo.xlsx')

    workbook.close
    compare_for_regression
  end
end
