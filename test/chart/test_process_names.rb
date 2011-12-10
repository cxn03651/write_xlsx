# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestProcessNames < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_process_names_both_nil
    expected = [nil, nil]
    result = @chart.__send__('process_names')
    assert_equal(expected, result)
  end

  def test_process_names_only_name
    expected = ['Text', nil]
    result = @chart.__send__('process_names', 'Text')
    assert_equal(expected, result)
  end

  def test_process_names_only_name_fomula_string
    expected = ['', '=Sheet1!$A$1']
    result = @chart.__send__('process_names', '=Sheet1!$A$1')
    assert_equal(expected, result)
  end
end
