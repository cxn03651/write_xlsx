# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteNumCache < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_num_cache
    expected = '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="5"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt></c:numCache>'
    @chart.__send__('write_num_cache', [1, 2, 3, 4, 5])
    result = @chart.instance_variable_get(:@writer).string
    assert_equal(expected, result)
  end
end
