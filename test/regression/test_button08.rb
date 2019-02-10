# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionButton08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_button08
    @xlsx = 'button08.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet1.insert_button('C2', {})

    worksheet2.write_comment('A1', 'Foo')

    worksheet2.comments_author = 'John'

    workbook.close
    compare_for_regression
  end
end
