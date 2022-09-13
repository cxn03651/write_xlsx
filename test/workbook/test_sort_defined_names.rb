# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx'
require 'stringio'

class TestSortDefinedNames < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
  end

  def test_sort_defined_names
    sorted = @workbook.__send__('sort_defined_names', unsorted)
    assert_equal(expected, sorted)
  end

  def test_extract_named_ranges
    sorted = @workbook.__send__('sort_defined_names', unsorted)
    assert_equal(@workbook.__send__('extract_named_ranges', sorted), named_ranges)
  end

  def unsorted
    [
      ["Bar",                    1, 'Sheet2!$A$1'],
      ["Bar",                    0, 'Sheet1!$A$1'],
      ["Abc",                   -1, 'Sheet1!$A$1'],
      ["Baz",                   -1, '0.98'],
      ["Bar",                    2, "'Sheet 3'!$A$1"],
      ["Foo",                   -1, 'Sheet1!$A$1'],
      ["Print_Titler",          -1, 'Sheet1!$A$1'],
      ["Print_Titlet",          -1, 'Sheet1!$A$1'],
      ["_Fog",                  -1, 'Sheet1!$A$1'],
      ["_Egg",                  -1, 'Sheet1!$A$1'],
      ["_xlnm.Print_Titles",     0, 'Sheet1!$1:$1'],
      ["_xlnm._FilterDatabase",  0, 'Sheet1!$G$1'],
      ["aaa",                    1, 'Sheet2!$A$1'],
      ["_xlnm.Print_Area",       0, 'Sheet1!$A$1:$H$10'],
      ["Car",                    2, '"Saab 900"']
    ]
  end

  def expected
    [
      ["_Egg",                  -1, 'Sheet1!$A$1'],
      ["_xlnm._FilterDatabase",  0, 'Sheet1!$G$1'],
      ["_Fog",                  -1, 'Sheet1!$A$1'],
      ["aaa",                    1, 'Sheet2!$A$1'],
      ["Abc",                   -1, 'Sheet1!$A$1'],
      ["Bar",                    2, "'Sheet 3'!$A$1"],
      ["Bar",                    0, 'Sheet1!$A$1'],
      ["Bar",                    1, 'Sheet2!$A$1'],
      ["Baz",                   -1, '0.98'],
      ["Car",                    2, '"Saab 900"'],
      ["Foo",                   -1, 'Sheet1!$A$1'],
      ["_xlnm.Print_Area",       0, 'Sheet1!$A$1:$H$10'],
      ["Print_Titler",          -1, 'Sheet1!$A$1'],
      ["_xlnm.Print_Titles",     0, 'Sheet1!$1:$1'],
      ["Print_Titlet",          -1, 'Sheet1!$A$1']
    ]
  end

  def named_ranges
    [
      '_Egg',
      '_Fog',
      'Sheet2!aaa',
      'Abc',
      "'Sheet 3'!Bar",
      'Sheet1!Bar',
      'Sheet2!Bar',
      'Foo',
      'Sheet1!Print_Area',
      'Print_Titler',
      'Sheet1!Print_Titles',
      'Print_Titlet'
    ]
  end
end
