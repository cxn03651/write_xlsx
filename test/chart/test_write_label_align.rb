# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteLabelAlign < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_label_align
    expected = '<c:lblAlgn val="ctr"/>'
    result = @chart.__send__('write_label_align', 'ctr')
    assert_equal(expected, result)
  end
end
