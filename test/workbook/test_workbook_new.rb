# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'stringio'

class TestWorkbookNew < Test::Unit::TestCase
  def test_workbook_new_without_param_raise
    assert_raise(ArgumentError) do
      WriteXLSX.new().close
    end
  end

  def test_workbook_new_with_null_string_raise
    assert_raise(RuntimeError) do
      WriteXLSX.new('').close
    end
  end
end
