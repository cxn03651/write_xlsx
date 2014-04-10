# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'

class TestColName < Test::Unit::TestCase
  def test_col_str_with_rational
    obj = ColName.instance
    col = Rational(10, 3)
    assert_nothing_raised do
      obj.col_str(col)
    end
  end
end
