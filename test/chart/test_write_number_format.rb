# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteNumberFormat < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_number_format
    axis = Writexlsx::Chart::Axis.new
    axis._num_format = 'General'
    axis._defaults   = { :num_format => 'General' }

    expected = '<c:numFmt formatCode="General" sourceLinked="1"/>'
    result = @chart.__send__('write_number_format', axis)
    assert_equal(expected, result)
  end

  def test_write_number_format02
    axis = Writexlsx::Chart::Axis.new
    axis._num_format = '#,##0.00'
    axis._defaults   = { :num_format => 'General' }

    expected = '<c:numFmt formatCode="#,##0.00" sourceLinked="0"/>'
    result = @chart.__send__('write_number_format', axis)
    assert_equal(expected, result)
  end

  def test_write_number_format03
    axis = Writexlsx::Chart::Axis.new
    axis._num_format = 'General'
    axis._defaults   = { :num_format => 'General' }

    expected = ''
    result = @chart.__send__('write_cat_number_format', axis)
    assert_equal(expected, result)
  end

  def test_write_number_format04
    axis = Writexlsx::Chart::Axis.new
    axis._num_format = '#,##0.00'
    axis._defaults   = { :num_format => 'General' }

    expected = '<c:numFmt formatCode="#,##0.00" sourceLinked="0"/>'
    result = @chart.__send__('write_cat_number_format', axis)
    assert_equal(expected, result)
  end
end
