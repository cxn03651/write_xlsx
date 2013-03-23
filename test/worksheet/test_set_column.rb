# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestSetColumn < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_set_column_with_level_minus
    level = -1
    @worksheet.set_column('B:G', nil, nil, 0, level)
    result = @worksheet.instance_variable_get(:@colinfo).first[5]
    assert_equal(0, result)
  end

  def test_set_column_with_level_8
    level = 8
    @worksheet.set_column('B:G', nil, nil, 0, level)
    result = @worksheet.instance_variable_get(:@colinfo).first[5]
    assert_equal(7, result)
  end
end
