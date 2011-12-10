# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestSortDefinedNames < Test::Unit::TestCase
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
      [ "Bar",                    1, %q(Sheet2!$A$1)       ],
      [ "Bar",                    0, %q(Sheet1!$A$1)       ],
      [ "Abc",                   -1, %q(Sheet1!$A$1)       ],
      [ "Baz",                   -1, %q(0.98)              ],
      [ "Bar",                    2, %q('Sheet 3'!$A$1)    ],
      [ "Foo",                   -1, %q(Sheet1!$A$1)       ],
      [ "Print_Titler",          -1, %q(Sheet1!$A$1)       ],
      [ "Print_Titlet",          -1, %q(Sheet1!$A$1)       ],
      [ "_Fog",                  -1, %q(Sheet1!$A$1)       ],
      [ "_Egg",                  -1, %q(Sheet1!$A$1)       ],
      [ "_xlnm.Print_Titles",     0, %q(Sheet1!$1:$1)      ],
      [ "_xlnm._FilterDatabase",  0, %q(Sheet1!$G$1)       ],
      [ "aaa",                    1, %q(Sheet2!$A$1)       ],
      [ "_xlnm.Print_Area",       0, %q(Sheet1!$A$1:$H$10) ],
      [ "Car",                    2, %q("Saab 900")        ]
    ]
  end

  def expected
    [
      [ "_Egg",                  -1, %q(Sheet1!$A$1)       ],
      [ "_xlnm._FilterDatabase",  0, %q(Sheet1!$G$1)       ],
      [ "_Fog",                  -1, %q(Sheet1!$A$1)       ],
      [ "aaa",                    1, %q(Sheet2!$A$1)       ],
      [ "Abc",                   -1, %q(Sheet1!$A$1)       ],
      [ "Bar",                    2, %q('Sheet 3'!$A$1)    ],
      [ "Bar",                    0, %q(Sheet1!$A$1)       ],
      [ "Bar",                    1, %q(Sheet2!$A$1)       ],
      [ "Baz",                   -1, %q(0.98)              ],
      [ "Car",                    2, %q("Saab 900")        ],
      [ "Foo",                   -1, %q(Sheet1!$A$1)       ],
      [ "_xlnm.Print_Area",       0, %q(Sheet1!$A$1:$H$10) ],
      [ "Print_Titler",          -1, %q(Sheet1!$A$1)       ],
      [ "_xlnm.Print_Titles",     0, %q(Sheet1!$1:$1)      ],
      [ "Print_Titlet",          -1, %q(Sheet1!$A$1)       ]
    ]
  end

  def named_ranges
    [
      %q(_Egg),
      %q(_Fog),
      %q(Sheet2!aaa),
      %q(Abc),
      %q('Sheet 3'!Bar),
      %q(Sheet1!Bar),
      %q(Sheet2!Bar),
      %q(Foo),
      %q(Sheet1!Print_Area),
      %q(Print_Titler),
      %q(Sheet1!Print_Titles),
      %q(Print_Titlet)
    ]
  end
end
