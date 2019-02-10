# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionCondFormat13 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_cond_format13
    @xlsx = 'cond_format04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # We manually set the indices to get the same order as the target file.
    format1 = workbook.add_format(:num_format => 2)
    format1.instance_variable_set(:@dxf_index, 1)

    format2 = workbook.add_format(:num_format => '0.000')
    format2.instance_variable_set(:@dxf_index, 0)

    worksheet.write('A1', 10)
    worksheet.write('A2', 20)
    worksheet.write('A3', 30)
    worksheet.write('A4', 40)

    options = {
      :type     => 'cell',
      :format   => format1,
      :criteria => '>',
      :value    => 2
    }

    worksheet.conditional_formatting('A1', options)

    # Test re-using options.
    options[:criteria] = '<'
    options[:value]    = 8
    options[:format]   = format2

    worksheet.conditional_formatting('A2', options)

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
