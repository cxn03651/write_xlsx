# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestDefinedName < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
  end

  def test_define_name_should_not_change_formula_string
    formula = '=0.98'
    @workbook.define_name('Name', formula)
    assert_equal('=0.98', formula)
  end
end
