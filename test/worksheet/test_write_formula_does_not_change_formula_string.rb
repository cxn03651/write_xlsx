# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteFormula < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_formula_does_not_change_formula_string
    formula = '=PI()'
    @worksheet.write('A1', formula)

    assert_equal('=PI()', formula)
  end
end
