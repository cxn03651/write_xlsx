# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestFormat < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @format   = @workbook.add_format
  end

  def test_set_align_with_frozen_parameter
    assert_nothing_raised do
      @format.set_align('LEFT'.freeze)
    end
  end
end
