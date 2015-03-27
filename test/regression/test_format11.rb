# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFormat11 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_format11
    @xlsx = 'format11.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    centered  = workbook.add_format(
      :align  => 'center',
      :valign => 'vcenter'
    )

    worksheet.write('B2', "Foo", centered)

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
