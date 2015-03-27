# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionLandscape01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_landscape01
    @xlsx = 'landscape01.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.write(0, 0, 'Foo')
    worksheet.set_landscape
    worksheet.paper = 9

    worksheet.vertical_dpi = 200

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
