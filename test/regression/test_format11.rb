# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFormat11 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format11
    @xlsx = 'format11.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    centered  = workbook.add_format(
      :align  => 'center',
      :valign => 'vcenter'
    )

    worksheet.write('B2', "Foo", centered)

    workbook.close
    compare_for_regression
  end
end
