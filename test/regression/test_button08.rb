# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionButton08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_button08
    @xlsx = 'button08.xlsx'
    workbook   = WriteXLSX.new(@xlsx)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet1.insert_button('C2', {})

    worksheet2.write_comment('A1', 'Foo')

    worksheet2.comments_author = 'John'

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
