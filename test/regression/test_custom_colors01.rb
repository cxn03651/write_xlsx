# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionCustomColors01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_custom_colors01
    @xlsx = 'custom_colors01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    workbook.set_custom_color(40, '#26DA55')
    workbook.set_custom_color(41, '#792DC8')
    workbook.set_custom_color(42, '#646462')

    color1 = workbook.add_format(:bg_color => 40)
    color2 = workbook.add_format(:bg_color => 41)
    color3 = workbook.add_format(:bg_color => 42)

    worksheet.write('A1', 'Foo', color1)
    worksheet.write('A2', 'Foo', color2)
    worksheet.write('A3', 'Foo', color3)

    workbook.close
    compare_for_regression
  end
end
